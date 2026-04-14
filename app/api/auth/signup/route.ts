import { NextResponse } from "next/server";

export const runtime = "nodejs";

export async function POST() {
  try {
    const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;

    if (!supabaseUrl) {
      return NextResponse.json({ error: "missing url env" }, { status: 500 });
    }

    const res = await fetch(`${supabaseUrl}/auth/v1/settings`, {
      method: "GET",
    });

    const text = await res.text();

    return new NextResponse(text, {
      status: res.status,
      headers: {
        "Content-Type": res.headers.get("content-type") || "application/json",
      },
    });
  } catch (err) {
    return NextResponse.json(
      {
        error: err instanceof Error ? err.message : "unknown error",
      },
      { status: 500 }
    );
  }
}