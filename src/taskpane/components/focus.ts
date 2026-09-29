// Moves focus to a screen's heading when the screen appears, so keyboard and screen reader users
// land at the top of the new content.
import { useEffect, useRef } from "react";

export function useFocusOnMount<T extends HTMLElement>(enabled = true) {
  const ref = useRef<T>(null);
  useEffect(() => {
    if (enabled) ref.current?.focus({ preventScroll: false });
  }, [enabled]);
  return ref;
}
