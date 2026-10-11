import {
  Component,
  type ComponentType,
  type ErrorInfo,
  type ReactNode,
} from 'react';

export interface ErrorFallbackProps {
  error: Error;
  resetError: () => void;
}

interface ErrorBoundaryProps {
  children: ReactNode;
  FallbackComponent?: ComponentType<ErrorFallbackProps>;
  /** Changing this clears a caught error. Pass the route to recover on navigation. */
  resetKey?: unknown;
}

interface ErrorBoundaryState {
  error: Error | null;
}

function toError(value: unknown): Error {
  if (value instanceof Error) {
    return value;
  }
  if (typeof value === 'string') {
    return new Error(value);
  }
  try {
    return new Error(JSON.stringify(value));
  } catch {
    return new Error(String(value));
  }
}

function DefaultFallback({ error, resetError }: ErrorFallbackProps) {
  return (
    <div
      className="min-h-screen w-full flex items-center justify-center p-6"
      style={{
        backgroundImage:
          'linear-gradient(180deg, #7ec8f5 0%, #b8e0f8 40%, #e8f4fc 70%, #f0e6c8 100%)',
      }}
    >
      <div className="pixel-panel max-w-lg w-full p-6 text-center">
        <div className="text-4xl mb-4 select-none">💥</div>

        <h1 className="pixel-font text-[13px] text-[#1a1a2e] mb-3 leading-relaxed">
          GAME OVER
        </h1>

        <p className="text-sm text-[#1a1a2e]/70 mb-4">
          This part of the app crashed. The rest is still running.
        </p>

        {import.meta.env.DEV ? (
          <pre className="mt-3 mb-5 overflow-x-auto rounded-none pixel-border-sm bg-[#fff8e7] p-3 text-left text-[11px] text-[#1a1a2e]">
            {error.message || String(error)}
          </pre>
        ) : null}

        <button
          type="button"
          onClick={resetError}
          className="pixel-btn bg-[#ffe566] px-6 py-2.5 text-[12px] font-bold text-[#1a1a2e] hover:bg-[#ffd700]"
        >
          TRY AGAIN
        </button>
      </div>
    </div>
  );
}

export class ErrorBoundary extends Component<
  ErrorBoundaryProps,
  ErrorBoundaryState
> {
  state: ErrorBoundaryState = { error: null };

  static getDerivedStateFromError(error: unknown): ErrorBoundaryState {
    return { error: toError(error) };
  }

  componentDidCatch(error: unknown, info: ErrorInfo): void {
    console.error(
      'ErrorBoundary caught an error:',
      toError(error),
      info.componentStack,
    );
  }

  componentDidUpdate(prevProps: ErrorBoundaryProps): void {
    if (
      this.state.error !== null &&
      prevProps.resetKey !== this.props.resetKey
    ) {
      this.resetError();
    }
  }

  resetError = (): void => {
    this.setState({ error: null });
  };

  render(): ReactNode {
    const { error } = this.state;
    if (error === null) {
      return this.props.children;
    }
    const Fallback = this.props.FallbackComponent ?? DefaultFallback;
    return <Fallback error={error} resetError={this.resetError} />;
  }
}
