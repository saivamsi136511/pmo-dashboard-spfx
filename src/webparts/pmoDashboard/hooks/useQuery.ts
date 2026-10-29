import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// ─── Custom useQuery (replaces @tanstack/react-query) ────────────────────────
// Written in TypeScript 3.3-compatible syntax (no optional chaining, no
// nullish coalescing, no const assertions, no template literal types).

export interface IUseQueryResult<TData, TError> {
  data: TData | undefined;
  isLoading: boolean;
  isError: boolean;
  error: TError | null;
  dataUpdatedAt: number;
  refetch: () => void;
}

export function useQuery<TData, TError>(
  queryFn: () => Promise<TData>
): IUseQueryResult<TData, TError> {
  const [data, setData] = React.useState<TData | undefined>(undefined);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [isError, setIsError] = React.useState<boolean>(false);
  const [error, setError] = React.useState<TError | null>(null);
  const [dataUpdatedAt, setDataUpdatedAt] = React.useState<number>(0);
  const [tick, setTick] = React.useState<number>(0);

  React.useEffect(() => {
    var cancelled = false;
    setIsLoading(true);
    setIsError(false);
    setError(null);

    queryFn()
      .then((result: TData) => {
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
  }, [tick]); // eslint-disable-line react-hooks/exhaustive-deps

  function refetch(): void {
    setTick((t: number) => t + 1);
  }

  return { data: data, isLoading: isLoading, isError: isError, error: error, dataUpdatedAt: dataUpdatedAt, refetch: refetch };
}

// ─── Props ───────────────────────────────────────────────────────────────────

export interface IDashboardViewProps {
  context: WebPartContext;
  title: string;
}
