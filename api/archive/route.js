import { NextResponse } from 'next/server';

export async function POST(request) {
  // Replace this with your actual Power Automate URL
  const POWER_AUTOMATE_URL = "https://default062a8e8e449048f39ee3b309e2cfa4.ad.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/36e5dea2ad0f4486ac1c61e45e6dde4d/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=vSvhWWzURBFmq2LDhc7ysp6wZ9blVACL2UNQ2SNRarA";

  try {
    const body = await request.json();

    const response = await fetch(POWER_AUTOMATE_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });

    // Capture the raw response to debug the 502 issue
    const data = await response.json();
    
    return NextResponse.json(data, { status: response.status });
  } catch (error) {
    return NextResponse.json({ error: 'Proxy failed', details: error.message }, { status: 500 });
  }
}
