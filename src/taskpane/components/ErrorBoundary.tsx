// Catches rendering errors and shows a plain message instead of a stack trace. Nothing is logged
// or sent anywhere (invariant 7).
import { Component, type ReactNode } from "react";

interface Props {
  children: ReactNode;
  /** Shown instead of the children after an error. Gets a function that tries again. */
  fallback: (retry: () => void) => ReactNode;
}

export class ErrorBoundary extends Component<Props, { failed: boolean }> {
  override state = { failed: false };

  static getDerivedStateFromError(): { failed: boolean } {
    return { failed: true };
  }

  override componentDidCatch(): void {
    // Deliberately empty: error details can quote workbook text, so they aren't kept or shown.
  }

  override render(): ReactNode {
    if (this.state.failed) return this.props.fallback(() => this.setState({ failed: false }));
    return this.props.children;
  }
}
