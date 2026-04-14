"use client";

import { useEffect, useState } from "react";
import { supabase } from "@/lib/supabase";
import { useRouter } from "next/navigation";

type Mode = "login" | "signup" | "forgot" | "reset";

export default function LoginPage() {
  const router = useRouter();

  const [mode, setMode] = useState<Mode>("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");
  const [errorMessage, setErrorMessage] = useState("");

  function clearStatus() {
    setMessage("");
    setErrorMessage("");
  }

  function changeMode(nextMode: Mode) {
    clearStatus();
    setPassword("");
    setConfirmPassword("");
    setMode(nextMode);
  }

  useEffect(() => {
    if (typeof window === "undefined") return;

    const hashParams = new URLSearchParams(window.location.hash.replace(/^#/, ""));
    const searchParams = new URLSearchParams(window.location.search);

    const type = hashParams.get("type") || searchParams.get("type");
    const error = hashParams.get("error") || searchParams.get("error");
    const errorCode = hashParams.get("error_code") || searchParams.get("error_code");
    const errorDescription =
      hashParams.get("error_description") || searchParams.get("error_description");
    const requestedMode = searchParams.get("mode");

    if (type === "recovery" || requestedMode === "reset") {
      setMode("reset");
      setMessage("Enter your new password below.");
    }

    if (error) {
      if (errorCode === "otp_expired") {
        setMode("forgot");
        setErrorMessage("This reset link expired. Enter your email and send a new reset link.");
      } else {
        setMode("forgot");
        setErrorMessage(
          decodeURIComponent(errorDescription || "Could not verify the reset link.")
        );
      }
    }

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((event) => {
      if (event === "PASSWORD_RECOVERY") {
        clearStatus();
        setMode("reset");
        setMessage("Recovery verified. Set your new password.");
      }
    });

    return () => {
      subscription.unsubscribe();
    };
  }, []);

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    clearStatus();
    setLoading(true);

    try {
      if (mode === "signup") {
        const { error } = await supabase.auth.signUp({
          email,
          password,
        });

        if (error) {
          setErrorMessage(error.message);
          return;
        }

        setMessage("Signup successful. Check your email if confirmation is enabled.");
        return;
      }

      if (mode === "login") {
        const { error } = await supabase.auth.signInWithPassword({
          email,
          password,
        });

        if (error) {
          setErrorMessage(error.message);
          return;
        }

        router.push("/");
        return;
      }

      if (mode === "forgot") {
        if (!email.trim()) {
          setErrorMessage("Enter your email first.");
          return;
        }

        const redirectTo =
          typeof window !== "undefined"
            ? `${window.location.origin}/login?mode=reset`
            : undefined;

        const { error } = await supabase.auth.resetPasswordForEmail(email, {
          redirectTo,
        });

        if (error) {
          setErrorMessage(error.message);
          return;
        }

        setMessage("Password reset email sent.");
        return;
      }

      if (mode === "reset") {
        if (!password || !confirmPassword) {
          setErrorMessage("Fill in both password fields.");
          return;
        }

        if (password !== confirmPassword) {
          setErrorMessage("Passwords do not match.");
          return;
        }

        const { error } = await supabase.auth.updateUser({
          password,
        });

        if (error) {
          setErrorMessage(error.message);
          return;
        }

        setMessage("Password updated successfully.");
        setMode("login");
        setPassword("");
        setConfirmPassword("");
        return;
      }
    } catch (err) {
      const message =
        err instanceof Error
          ? err.message
          : "Network error. Check Supabase URL/key and Auth URL config.";
      setErrorMessage(message);
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="min-h-screen flex items-center justify-center bg-white px-6">
      <div className="w-full max-w-md rounded-3xl border border-gray-200 p-8 shadow-sm">
        <h1 className="text-3xl font-semibold tracking-tight">OnTrack</h1>

        <p className="mt-2 text-sm text-gray-600">
          {mode === "login" && "Sign in to your account"}
          {mode === "signup" && "Create your account"}
          {mode === "forgot" && "Send yourself a password reset email"}
          {mode === "reset" && "Choose a new password"}
        </p>

        {message ? (
          <div className="mt-4 rounded-xl border border-green-200 bg-green-50 px-4 py-3 text-sm text-green-700">
            {message}
          </div>
        ) : null}

        {errorMessage ? (
          <div className="mt-4 rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
            {errorMessage}
          </div>
        ) : null}

        <form onSubmit={handleSubmit} className="mt-6 space-y-4">
          {mode !== "reset" && (
            <div>
              <label className="mb-1 block text-sm font-medium">Email</label>
              <input
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                required
                className="w-full rounded-xl border border-gray-300 px-4 py-3 outline-none focus:border-black"
              />
            </div>
          )}

          {mode !== "forgot" && (
            <div>
              <label className="mb-1 block text-sm font-medium">
                {mode === "reset" ? "New password" : "Password"}
              </label>
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                required
                className="w-full rounded-xl border border-gray-300 px-4 py-3 outline-none focus:border-black"
              />
            </div>
          )}

          {mode === "reset" && (
            <div>
              <label className="mb-1 block text-sm font-medium">Confirm new password</label>
              <input
                type="password"
                value={confirmPassword}
                onChange={(e) => setConfirmPassword(e.target.value)}
                required
                className="w-full rounded-xl border border-gray-300 px-4 py-3 outline-none focus:border-black"
              />
            </div>
          )}

          <button
            type="submit"
            disabled={loading}
            className="w-full rounded-xl bg-black px-4 py-3 text-white disabled:opacity-60"
          >
            {loading
              ? "Please wait..."
              : mode === "login"
              ? "Log in"
              : mode === "signup"
              ? "Sign up"
              : mode === "forgot"
              ? "Send reset email"
              : "Save new password"}
          </button>
        </form>

        <div className="mt-4 flex flex-col gap-2 text-sm text-gray-600">
          {mode === "login" && (
            <>
              <button
                type="button"
                onClick={() => changeMode("forgot")}
                className="text-left underline"
              >
                Forgot password?
              </button>

              <button
                type="button"
                onClick={() => changeMode("signup")}
                className="text-left underline"
              >
                Need an account? Sign up
              </button>
            </>
          )}

          {mode === "signup" && (
            <button
              type="button"
              onClick={() => changeMode("login")}
              className="text-left underline"
            >
              Already have an account? Log in
            </button>
          )}

          {mode === "forgot" && (
            <button
              type="button"
              onClick={() => changeMode("login")}
              className="text-left underline"
            >
              Back to login
            </button>
          )}

          {mode === "reset" && (
            <button
              type="button"
              onClick={() => changeMode("login")}
              className="text-left underline"
            >
              Back to login
            </button>
          )}
        </div>
      </div>
    </main>
  );
}