
import { GDRIVE_FOLDER_ID } from '../constants';

export const config = {
  runtime: 'edge',
};

function cleanCredential(val: string | undefined): string {
  if (!val) return '';
  return val.trim().replace(/^["']|["']$/g, '').replace(/\\n/g, '\n');
}

async function getAccessToken(clientId: string, clientSecret: string, refreshToken: string) {
  const params = new URLSearchParams();
  params.append('client_id', clientId);
  params.append('client_secret', clientSecret);
  params.append('refresh_token', refreshToken);
  params.append('grant_type', 'refresh_token');

  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  });

  const responseText = await res.text();
  if (!res.ok) {
    throw new Error(`Google Auth Error: ${responseText}`);
  }
  
  const data = JSON.parse(responseText);
  return data.access_token;
}

export default async function handler(req: Request) {
  if (req.method !== 'DELETE' && req.method !== 'POST') {
    return new Response(JSON.stringify({ error: 'Method not allowed' }), { status: 405 });
  }

  try {
    const url = new URL(req.url);
    const fileId = url.searchParams.get('fileId');

    if (!fileId) {
      return new Response(JSON.stringify({ error: 'fileId is required' }), { status: 400 });
    }

    const clientId = cleanCredential(process.env.GOOGLE_CLIENT_ID || process.env.VITE_GOOGLE_CLIENT_ID);
    const clientSecret = cleanCredential(process.env.GOOGLE_CLIENT_SECRET || process.env.VITE_GOOGLE_CLIENT_SECRET);
    const refreshToken = cleanCredential(process.env.GOOGLE_REFRESH_TOKEN || process.env.VITE_GOOGLE_REFRESH_TOKEN);

    if (!clientId || !clientSecret || !refreshToken) {
      return new Response(JSON.stringify({ error: 'Missing Google credentials' }), { status: 500 });
    }

    const accessToken = await getAccessToken(clientId, clientSecret, refreshToken);

    // Permanent delete using files.delete
    const driveResponse = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}`, {
      method: 'DELETE',
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (!driveResponse.ok && driveResponse.status !== 404) {
      const errorData = await driveResponse.text();
      console.error('GOOGLE_DRIVE_DELETE_ERROR:', errorData);
      return new Response(JSON.stringify({ error: 'Gagal menghapus file dari Google Drive', detail: errorData }), { status: driveResponse.status });
    }

    return new Response(JSON.stringify({ success: true }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    console.error('DELETE_HANDLER_ERROR:', error);
    return new Response(JSON.stringify({ error: error instanceof Error ? error.message : 'Internal Server Error' }), { status: 500 });
  }
}
