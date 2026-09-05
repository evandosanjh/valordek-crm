import type { Config } from '@netlify/functions'

const SUPABASE_URL = process.env.SUPABASE_URL || ''
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || ''
const MS_TENANT_ID = process.env.MS_TENANT_ID || ''
const MS_CLIENT_ID = process.env.MS_CLIENT_ID || ''
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET || ''
const MAILBOX = 'info@valordek.com'

async function graphToken() {
  const body = new URLSearchParams({
    client_id: MS_CLIENT_ID,
    client_secret: MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  })
  const r = await fetch(`https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body,
  })
  if (!r.ok) throw new Error(`Microsoft token failed: ${r.status}`)
  return (await r.json()).access_token as string
}

async function sb(path: string, init: RequestInit = {}) {
  return fetch(`${SUPABASE_URL}${path}`, {
    ...init,
    headers: {
      apikey: SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
      'Content-Type': 'application/json',
      ...(init.headers || {}),
    },
  })
}

function looksLikeValordekEstimate(text: string) {
  const t = text.toLowerCase()
  if (!t.includes('valordek')) return false
  if (!t.includes('estimate')) return false
  return true
}

export default async () => {
  if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY || !MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) {
    console.log('Background sync skipped: required environment variables are missing.')
    return
  }

  const token = await graphToken()
  const graph = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(MAILBOX)}/messages?$top=25&$orderby=receivedDateTime%20desc&$select=id,subject,receivedDateTime,from,bodyPreview,hasAttachments`
  const r = await fetch(graph, { headers: { Authorization: `Bearer ${token}` } })
  if (!r.ok) throw new Error(`Graph messages failed: ${r.status}`)
  const messages = (await r.json()).value || []

  for (const m of messages) {
    const combined = `${m.subject || ''}\n${m.bodyPreview || ''}`
    if (!looksLikeValordekEstimate(combined)) continue

    const exists = await sb(`/rest/v1/estimate_imports?provider=eq.microsoft&provider_message_id=eq.${encodeURIComponent(m.id)}&select=id&limit=1`)
    if (!exists.ok) continue
    if ((await exists.json()).length) continue

    await sb('/rest/v1/estimate_imports', {
      method: 'POST',
      headers: { Prefer: 'return=minimal' },
      body: JSON.stringify({
        owner_id: process.env.VALORDEK_OWNER_ID,
        provider: 'microsoft',
        provider_message_id: m.id,
        subject: m.subject || null,
        sender_email: m.from?.emailAddress?.address || null,
        received_at: m.receivedDateTime || null,
        parse_status: 'pending',
        confidence: null,
        extracted_data: {
          source: 'background_sync',
          body_preview: m.bodyPreview || '',
          has_attachments: !!m.hasAttachments,
        },
      }),
    })
  }

  await sb(`/rest/v1/mailbox_connections?owner_id=eq.${encodeURIComponent(process.env.VALORDEK_OWNER_ID || '')}`, {
    method: 'PATCH',
    body: JSON.stringify({ status: 'connected', last_sync_at: new Date().toISOString(), last_error: null, updated_at: new Date().toISOString() }),
  })
}

export const config: Config = { schedule: '*/15 * * * *' }
