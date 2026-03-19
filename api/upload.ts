
import { GDRIVE_FOLDER_ID } from '../constants';

export const config = {
  runtime: 'edge',
};

function cleanCredential(val: string | undefined): string {
  if (!val) return '';
  // Membersihkan spasi, karakter baris baru, dan tanda kutip yang sering terbawa saat copy-paste dari dashboard Vercel
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
  let data;
  try {
    data = JSON.parse(responseText);
  } catch (e) {
    data = { error: 'Invalid JSON Response', raw: responseText };
  }
  
  if (!res.ok) {
    console.error(`GOOGLE_OAUTH_ERROR [${res.status}]:`, responseText);
    let customMessage = data.error_description || data.error || 'Gagal autentikasi Google';
    
    // Memberikan informasi lebih spesifik jika token tidak valid
    if (data.error === 'invalid_grant') {
      customMessage = 'Refresh Token tidak valid/kadaluarsa. Harap periksa kembali variabel GOOGLE_REFRESH_TOKEN di Vercel.';
    }
    
    // Melempar error dengan status asli agar handler bisa merespon dengan status yang sama
    const error: any = new Error(customMessage);
    error.status = res.status;
    throw error;
  }
  
  return data.access_token;
}

export default async function handler(req: Request) {
  if (req.method !== 'POST') {
    return new Response(JSON.stringify({ error: 'Method not allowed' }), { status: 405 });
  }

  try {
    const clientId = cleanCredential(process.env.GOOGLE_CLIENT_ID || process.env.VITE_GOOGLE_CLIENT_ID);
    const clientSecret = cleanCredential(process.env.GOOGLE_CLIENT_SECRET || process.env.VITE_GOOGLE_CLIENT_SECRET);
    const refreshToken = cleanCredential(process.env.GOOGLE_REFRESH_TOKEN || process.env.VITE_GOOGLE_REFRESH_TOKEN);

    if (!clientId || !clientSecret || !refreshToken) {
      return new Response(JSON.stringify({ 
        error: 'Kredensial OAuth2 belum lengkap di Vercel. Pastikan CLIENT_ID, SECRET, dan REFRESH_TOKEN sudah diatur.' 
      }), { status: 500 });
    }

    const formData = await req.formData();
    const file = formData.get('file') as File;

    if (!file) {
      return new Response(JSON.stringify({ error: 'Tidak ada file yang dikirim' }), { status: 400 });
    }

    // Mendapatkan access token dengan penanganan error status
    let accessToken;
    try {
      accessToken = await getAccessToken(clientId, clientSecret, refreshToken);
    } catch (authError: any) {
      return new Response(JSON.stringify({ error: authError.message }), { 
        status: authError.status || 401 
      });
    }

    const metadata = {
      name: `HUREMA_${Date.now()}_${file.name}`,
      parents: [GDRIVE_FOLDER_ID],
      mimeType: file.type,
    };

    const boundary = 'hurema_upload_boundary';
    const metadataPart = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n`;
    const mediaPart = `--${boundary}\r\nContent-Type: ${file.type}\r\n\r\n`;
    const closingPart = `\r\n--${boundary}--`;

    const fileBuffer = await file.arrayBuffer();
    const bodyArr = new Uint8Array([
      ...new TextEncoder().encode(metadataPart),
      ...new TextEncoder().encode(mediaPart),
      ...new Uint8Array(fileBuffer),
      ...new TextEncoder().encode(closingPart),
    ]);

    const driveResponse = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': `multipart/related; boundary=${boundary}`,
      },
      body: bodyArr,
    });

    const driveData = await driveResponse.json();

    if (!driveResponse.ok) {
      console.error('GOOGLE_DRIVE_API_ERROR:', driveData);
      return new Response(JSON.stringify({ 
        error: 'Gagal mengunggah ke Google Drive', 
        detail: driveData 
      }), { status: driveResponse.status });
    }

    return new Response(JSON.stringify({ id: driveData.id }), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    console.error('CRITICAL_UPLOAD_HANDLER_ERROR:', error);
    return new Response(JSON.stringify({ 
      error: error instanceof Error ? error.message : 'Internal Server Error' 
    }), { status: 500 });
  }
}
