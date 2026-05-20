import { NextResponse } from 'next/server';

const corsHeaders = {
  'Access-Control-Allow-Origin': 'https://zachariasd.github.io',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
};

export async function OPTIONS() {
  return new NextResponse(null, {
    status: 200,
    headers: corsHeaders,
  });
}

export async function POST(request) {
  const POWER_AUTOMATE_URL = "YOUR_POWER_AUTOMATE_URL";

  try {
    const body = await request.json();
    
    const response = await fetch(POWER_AUTOMATE_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });

    const data = await response.json();
    
    return NextResponse.json(data, { 
      status: response.status,
      headers: corsHeaders 
    });
  } catch (error) {
    console.error("Proxy Crash Details:", error);
    return NextResponse.json({ error: 'Proxy failed', details: error.message }, { 
      status: 500,
      headers: corsHeaders 
    });
  }
}
