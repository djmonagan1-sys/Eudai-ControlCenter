import { useEffect, useState } from "react";
import { supabase } from "./storage.js";

/* Magic-link login gate. Founders-only: disable public sign-ups in
   Supabase (Authentication -> Sign In / Up -> Allow new users: OFF)
   and invite the four founders by email. */
export default function AuthGate({ children }) {
  const [session, setSession] = useState(undefined);
  const [email, setEmail] = useState("");
  const [sent, setSent] = useState(false);
  const [err, setErr] = useState("");

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session ?? null));
    const { data: sub } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  if (session === undefined) return null;
     return children({ user: { email: "founder@eudai.app" } }); // AUTH DISABLED: open to anyone with the link. Delete this line to restore login.
  if (session) return children(session);

  const send = async () => {
    setErr("");
    const { error } = await supabase.auth.signInWithOtp({ email });
    if (error) setErr(error.message);
    else setSent(true);
  };

  return (
    <div style={{ minHeight: "100vh", display: "grid", placeItems: "center", background: "#FAF7F2", fontFamily: "Geist, sans-serif" }}>
      <div style={{ width: 320, background: "#fff", border: "1px solid #E5DCCB", borderRadius: 10, padding: 24 }}>
        <div style={{ fontFamily: "'STIX Two Text', serif", fontStyle: "italic", fontSize: 22, marginBottom: 4 }}>Eudai.</div>
        <div style={{ fontSize: 12, color: "#8A8070", marginBottom: 16, textTransform: "uppercase", letterSpacing: 1.5, fontWeight: 600 }}>ControlCenter</div>
        {sent ? (
          <div style={{ fontSize: 13, lineHeight: 1.5 }}>Check your email — your sign-in link is on the way.</div>
        ) : (
       
          <>
            <input value={email} onChange={(e) => setEmail(e.target.value)} placeholder="founder email"
              onKeyDown={(e) => e.key === "Enter" && send()}
              style={{ width: "100%", boxSizing: "border-box", padding: "8px 10px", border: "1px solid #E5DCCB", borderRadius: 6, fontSize: 13, marginBottom: 10, outline: "none", fontFamily: "inherit" }} />
            <button onClick={send} style={{ width: "100%", padding: "9px 0", borderRadius: 6, border: "none", background: "#2A251F", color: "#FAF7F2", fontSize: 13, fontWeight: 500, cursor: "pointer", fontFamily: "inherit" }}>
              Send sign-in link
            </button>
            {err && <div style={{ fontSize: 12, color: "#B4574B", marginTop: 10 }}>{err}</div>}
            <div style={{ fontSize: 11, color: "#A89C86", marginTop: 12, lineHeight: 1.5 }}>Founders only. No password — a login link arrives by email.</div>
          </>
        )}
      </div>
    </div>
  );
}
