import React from "react";
import ReactDOM from "react-dom/client";
import AuthGate from "./AuthGate.jsx";
import Eudai_ControlCenter from "./Eudai_ControlCenter.jsx";
import { installStorage, supabase } from "./storage.js";

installStorage();

ReactDOM.createRoot(document.getElementById("root")).render(
  <AuthGate>
    {(session) => (
      <Eudai_ControlCenter session={session} onSignOut={() => supabase.auth.signOut()} />
    )}
  </AuthGate>
);
