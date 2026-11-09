import { Dispatch, SetStateAction } from "react";

export function toggleTheme(dark: Boolean, setDark: Dispatch<SetStateAction<boolean>>) {
  const next = !dark;
  setDark(next);
  document.documentElement.classList.toggle("dark", next);
  try {
    localStorage.setItem("helpdesk-theme", next ? "dark" : "light");
  } catch {}
}
