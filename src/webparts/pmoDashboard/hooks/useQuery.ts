import * as React from 'react';

// ─── Lightweight useQuery replacement ────────────────────────────────────────
//
// @tanstack/react-query@4 uses TypeScript 4+ syntax in its .d.ts files which
// is incompatible with SPFx 1.11's TypeScript 3.3 compiler.
// This module provides the same API surface used in this project.

export interface UseQueryResult<TData, TError = Error> {
  data: TData | undefined;
  isLoading: boolean;
  isError: boolean;
  error: TError | null;
  dataUpdatedAt: number;
  refetch: () => void;
}

export function useQuery<TData, TError = Error>(
  _queryKey: unknown[],
  queryFn: () => Promise<TData>
): UseQueryResult<TData, TError> {
  const [data, setData] = React.useState<TData | undefined>(undefined);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [isError, setIsError] = React.useState<boolean>(false);
  const [error, setError] = React.useState<TError | null>(null);
  const [dataUpdatedAt, setDataUpdatedAt] = React.useState<number>(0);
  const [tick, setTick] = React.useState<number>(0);

  React.useEffect(() => {
    let cancelled = false;
    setIsLoading(true);
    setIsError(false);
    setError(null);

    queryFn()
      .then((result) => {
        if (!cancelled) {
          setData(result);
          setDataUpdatedAt(Date.now());
          setIsLoading(false);
        }
      })
      .catch((err: TError) => {
        if (!cancelled) {
          setIsError(true);
          setError(err);
          setIsLoading(false);
        }
      });

    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tick]);

  const refetch = React.useCallback(() => setTick((t) => t + 1), []);

  return { data, isLoading, isError, error, dataUpdatedAt, refetch };
}
