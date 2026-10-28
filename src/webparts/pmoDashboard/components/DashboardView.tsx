import * as React from 'react';
import { useQuery } from '@tanstack/react-query';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { SharePointListService, IProject } from '../services/SharePointListService';

// ─── Props ────────────────────────────────────────────────────────────────────

export interface IDashboardViewProps {
  context: WebPartContext;
  title: string;
}

// ─── Status Badge Configuration ──────────────────────────────────────────────

interface IBadgeStyle {
  background: string;
  color: string;
  border: string;
}

const STATUS_BADGE_MAP: Record<string, IBadgeStyle> = {
  Active:    { background: '#d1fae5', color: '#065f46', border: '1px solid #6ee7b7' },
  'On Hold': { background: '#fef3c7', color: '#92400e', border: '1px solid #fcd34d' },
  Completed: { background: '#dbeafe', color: '#1e3a8a', border: '1px solid #93c5fd' },
  Cancelled: { background: '#fee2e2', color: '#991b1b', border: '1px solid #fca5a5' },
};

const DEFAULT_BADGE: IBadgeStyle = {
  background: '#f3f4f6',
  color: '#374151',
  border: '1px solid #d1d5db',
};

function getStatusBadge(status: string): IBadgeStyle {
  return STATUS_BADGE_MAP[status] ?? DEFAULT_BADGE;
}

// ─── Inline Style Constants ───────────────────────────────────────────────────

const styles = {
  wrapper: {
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
    padding: '24px',
    background: '#ffffff',
    minHeight: '200px',
  } as React.CSSProperties,

  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '20px',
  } as React.CSSProperties,

  title: {
    fontSize: '20px',
    fontWeight: 600,
    color: '#111827',
    margin: 0,
  } as React.CSSProperties,

  countBadge: {
    fontSize: '12px',
    fontWeight: 500,
    background: '#eff6ff',
    color: '#1d4ed8',
    border: '1px solid #bfdbfe',
    borderRadius: '12px',
    padding: '2px 10px',
  } as React.CSSProperties,

  table: {
    width: '100%',
    borderCollapse: 'collapse' as const,
    fontSize: '14px',
  } as React.CSSProperties,

  th: {
    textAlign: 'left' as const,
    padding: '10px 14px',
    background: '#f9fafb',
    color: '#6b7280',
    fontWeight: 600,
    fontSize: '12px',
    textTransform: 'uppercase' as const,
    letterSpacing: '0.05em',
    borderBottom: '2px solid #e5e7eb',
  } as React.CSSProperties,

  td: {
    padding: '12px 14px',
    borderBottom: '1px solid #f3f4f6',
    color: '#374151',
    verticalAlign: 'middle' as const,
  } as React.CSSProperties,

  badge: (style: IBadgeStyle): React.CSSProperties => ({
    display: 'inline-block',
    padding: '3px 10px',
    borderRadius: '12px',
    fontSize: '12px',
    fontWeight: 600,
    ...style,
  }),

  managerCell: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  } as React.CSSProperties,

  avatar: (initials: string): React.CSSProperties => ({
    width: '28px',
    height: '28px',
    borderRadius: '50%',
    background: '#4f46e5',
    color: '#fff',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '11px',
    fontWeight: 700,
    flexShrink: 0,
  }),

  stateBox: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexDirection: 'column' as const,
    gap: '8px',
    padding: '48px 0',
    color: '#6b7280',
  } as React.CSSProperties,

  spinner: {
    width: '36px',
    height: '36px',
    border: '3px solid #e5e7eb',
    borderTop: '3px solid #4f46e5',
    borderRadius: '50%',
    animation: 'pmo-spin 0.8s linear infinite',
  } as React.CSSProperties,

  errorBox: {
    background: '#fff1f2',
    border: '1px solid #fecdd3',
    borderRadius: '8px',
    padding: '16px 20px',
    color: '#9f1239',
    fontSize: '14px',
  } as React.CSSProperties,

  emptyText: {
    fontSize: '14px',
  } as React.CSSProperties,
} as const;

// ─── Helper: derive initials from a display name ──────────────────────────────

function toInitials(name: string): string {
  return name
    .split(' ')
    .filter(Boolean)
    .slice(0, 2)
    .map((w) => w[0].toUpperCase())
    .join('');
}

// ─── Main Component ───────────────────────────────────────────────────────────

const DashboardView: React.FC<IDashboardViewProps> = ({ context, title }) => {
  // Instantiate the service. Memoized so the object reference is stable across
  // re-renders (context reference is stable for the lifetime of the web part).
  const service = React.useMemo(
    () => new SharePointListService(context),
    [context]
  );

  const {
    data: projects,
    isLoading,
    isError,
    error,
    dataUpdatedAt,
  } = useQuery<IProject[], Error>(
    // Cache key — scoped to the web's absolute URL so multi-tenant scenarios
    // never mix data across different sites.
    ['pmo-projects', context.pageContext.web.absoluteUrl],
    () => service.getProjects(),
    {
      // Inherit defaults from QueryClient but allow per-query overrides here if needed.
    }
  );

  // Inject keyframe animation once (no CSS-in-JS library required)
  React.useEffect(() => {
    const styleId = 'pmo-dashboard-keyframes';
    if (!document.getElementById(styleId)) {
      const styleEl = document.createElement('style');
      styleEl.id = styleId;
      styleEl.textContent = `@keyframes pmo-spin { to { transform: rotate(360deg); } }`;
      document.head.appendChild(styleEl);
    }
  }, []);

  // ── Render: Loading ──────────────────────────────────────────────────────
  if (isLoading) {
    return (
      <div style={styles.wrapper}>
        <div style={styles.stateBox}>
          <div style={styles.spinner} role="progressbar" aria-label="Loading projects" />
          <span>Loading projects…</span>
        </div>
      </div>
    );
  }

  // ── Render: Error ────────────────────────────────────────────────────────
  if (isError) {
    const errorMessage = error?.message ?? 'An unknown error occurred.';
    return (
      <div style={styles.wrapper}>
        <div style={styles.errorBox} role="alert">
          <strong>Failed to load PMO Projects</strong>
          <br />
          {errorMessage}
        </div>
      </div>
    );
  }

  const projectList: IProject[] = projects ?? [];
  const updatedLabel = dataUpdatedAt
    ? new Date(dataUpdatedAt).toLocaleTimeString()
    : '';

  // ── Render: Table ────────────────────────────────────────────────────────
  return (
    <div style={styles.wrapper}>
      {/* Keyframe injection also happens via useEffect, but the style tag is
          harmless here for fallback clarity in older browsers. */}

      <div style={styles.header}>
        <h2 style={styles.title}>{title}</h2>
        {projectList.length > 0 && (
          <span style={styles.countBadge}>
            {projectList.length} project{projectList.length !== 1 ? 's' : ''}
          </span>
        )}
      </div>

      {projectList.length === 0 ? (
        <div style={styles.stateBox}>
          <span role="img" aria-label="No data">📋</span>
          <span style={styles.emptyText}>No projects found in PMO_Projects list.</span>
        </div>
      ) : (
        <>
          <table style={styles.table} role="table" aria-label="PMO Projects">
            <thead>
              <tr>
                <th style={styles.th}>#</th>
                <th style={styles.th}>Project Name</th>
                <th style={styles.th}>Status</th>
                <th style={styles.th}>Project Manager</th>
              </tr>
            </thead>
            <tbody>
              {projectList.map((project, index) => {
                const badgeStyle = getStatusBadge(project.Project_Status);
                const managerName = project.Project_Manager?.Title ?? '—';
                const initials = managerName !== '—' ? toInitials(managerName) : '?';

                return (
                  <tr
                    key={project.Id}
                    style={{
                      background: index % 2 === 0 ? '#ffffff' : '#fafafa',
                      transition: 'background 0.15s',
                    }}
                    onMouseEnter={(e) => {
                      (e.currentTarget as HTMLTableRowElement).style.background = '#eff6ff';
                    }}
                    onMouseLeave={(e) => {
                      (e.currentTarget as HTMLTableRowElement).style.background =
                        index % 2 === 0 ? '#ffffff' : '#fafafa';
                    }}
                  >
                    <td style={{ ...styles.td, color: '#9ca3af', width: '40px' }}>
                      {index + 1}
                    </td>
                    <td style={{ ...styles.td, fontWeight: 500, color: '#111827' }}>
                      {project.Title}
                    </td>
                    <td style={styles.td}>
                      <span style={styles.badge(badgeStyle)}>
                        {project.Project_Status}
                      </span>
                    </td>
                    <td style={styles.td}>
                      <div style={styles.managerCell}>
                        <div
                          style={styles.avatar(initials)}
                          aria-hidden="true"
                          title={managerName}
                        >
                          {initials}
                        </div>
                        <span>{managerName}</span>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          {updatedLabel && (
            <p
              style={{
                margin: '12px 0 0',
                fontSize: '11px',
                color: '#9ca3af',
                textAlign: 'right',
              }}
            >
              Last refreshed at {updatedLabel}
            </p>
          )}
        </>
      )}
    </div>
  );
};

export default DashboardView;
