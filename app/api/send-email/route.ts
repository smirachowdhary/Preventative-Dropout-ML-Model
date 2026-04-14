import { NextResponse } from "next/server";
import { Resend } from "resend";

export const runtime = "nodejs";

const resend = new Resend(process.env.RESEND_API_KEY);

export async function POST(req: Request) {
  try {
    const { email } = await req.json();

    const data = await resend.emails.send({
      from: "onboarding@resend.dev", // works for testing
      to: email,
      subject: "Welcome to OnTrack 🎉",
      html: `
        <h2>Welcome to OnTrack!</h2>
        <p>Your account is ready 🚀</p>
      `,
    });

    return NextResponse.json({ ok: true, data });
  } catch (err) {
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Email failed" },
      { status: 500 }
    );
  }
}