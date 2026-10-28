import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { SharePointListService, IProject } from '../services/SharePointListService';
import { useQuery } from '../hooks/useQuery';

export interface IDashboardViewProps {
  context: WebPartContext;
  title: string;
}

// ─── Status Badge Styles ─────────────────────────────────────────────────────

interface IBadgeStyle {
  background: string;
  color: string;
  border: string;
}

const STATUS_BADGE_MAP: { [key: string]: IBadgeStyle } = {
  'Active':    { background: '#d1fae5', color: '#065f46', border: '1px solid #6ee7b7' },
  'On Hold':   { background: '#fef3c7', color: '#92400e', border: '1px solid #fcd34d' },
  'Completed': { background: '#dbeafe', color: '#1e3a8a', border: '1px solid #93c5fd' },
  'Cancelled': { background: '#fee2e2', color: '#991b1b', border: '1px solid #fca5a5' },
};

const DEFAULT_BADGE: IBadgeStyle = {
  background: '#f3f4f6',
  color: '#374151',
  border: '1px solid #d1d5db',
};

function getStatusBadge(status: string): IBadgeStyle {
  return STATUS_BADGE_MAP[status] || DEFAULT_BADGE;
}

function toInitials(name: string): string {
  return name
    .split(' ')
    .filter(Boolean)
    .slice(0, 2)
    .map(function(w) { return w[0].toUpperCase(); })
    .join('');
}

// ─── Styles ──────────────────────────────────────────────────────────────────

const wrapperStyle: React.CSSProperties = {
  fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  padding: '24px',
  background: '#ffffff',
  minHeight: '200px',
};

const headerStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  marginBottom: '20px',
};

const titleStyle: React.CSSProperties = {
  fontSize: '20px',
  fontWeight: 600,
  color: '#111827',
  margin: 0,
};

const countBadgeStyle: React.CSSProperties = {
  fontSize: '12px',
  fontWeight: 500,
  background: '#eff6ff',
  color: '#1d4ed8',
  border: '1px solid #bfdbfe',
  borderRadius: '12px',
  padding: '2px 10px',
};

const tableStyle: React.CSSProperties = {
  width: '100%',
  borderCollapse: 'collapse',
  fontSize: '14px',
};

const thStyle: React.CSSProperties = {
  textAlign: 'left',
  padding: '10px 14px',
  background: '#f9fafb',
  color: '#6b7280',
  fontWeight: 600,
  fontSize: '12px',
  textTransform: 'uppercase',
  letterSpacing: '0.05em',
  borderBottom: '2px solid #e5e7eb',
};

const tdStyle: React.CSSProperties = {
  padding: '12px 14px',
  borderBottom: '1px solid #f3f4f6',
  color: '#374151',
  verticalAlign: 'middle',
};

const stateBoxStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  flexDirection: 'column',
  gap: '8px',
  padding: '48px 0',
  color: '#6b7280',
};

const spinnerStyle: React.CSSProperties = {
  width: '36px',
  height: '36px',
  border: '3px solid #e5e7eb',
  borderTop: '3px solid #4f46e5',
  borderRadius: '50%',
  animation: 'pmo-spin 0.8s linear infinite',
};

const errorBoxStyle: React.CSSProperties = {
  background: '#fff1f2',
  border: '1px solid #fecdd3',
  borderRadius: '8px',
  padding: '16px 20px',
  color: '#9f1239',
  fontSize: '14px',
};

// ─── Component ────────────────────────────────────────────────────────────────

const DashboardView: React.FC<IDashboardViewProps> = ({ context, title }) => {
  const service = React.useMemo(
    () => new SharePointListService(context),
    [context]
  );

  const { data: projects, isLoading, isError, error, dataUpdatedAt } =
    useQuery<IProject[], Error>(
      ['pmo-projects', context.pageContext.web.absoluteUrl],
      () => service.getProjects()
    );

  // Inject keyframe animation once
  React.useEffect(() => {
    const styleId = 'pmo-dashboard-keyframes';
    if (!document.getElementById(styleId)) {
      const styleEl = document.createElement('style');
      styleEl.id = styleId;
      styleEl.textContent = '@keyframes pmo-spin { to { transform: rotate(360deg); } }';
      document.head.appendChild(styleEl);
    }
  }, []);

  if (isLoading) {
    return (
      <div style={wrapperStyle}>
        <div style={stateBoxStyle}>
          <div style={spinnerStyle} />
          <span>Loading projects…</span>
        </div>
      </div>
    );
  }

  if (isError) {
    return (
      <div style={wrapperStyle}>
        <div style={errorBoxStyle}>
          <strong>Failed to load PMO Projects</strong><br />
          {error ? error.message : 'An unknown error occurred.'}
        </div>
      </div>
    );
  }

  const projectList: IProject[] = projects || [];
  const updatedLabel = dataUpdatedAt ? new Date(dataUpdatedAt).toLocaleTimeString() : '';

  return (
    <div style={wrapperStyle}>
      <div style={headerStyle}>
        <h2 style={titleStyle}>{title}</h2>
        {projectList.length > 0 && (
          <span style={countBadgeStyle}>
            {projectList.length} project{projectList.length !== 1 ? 's' : ''}
          </span>
        )}
      </div>

      {projectList.length === 0 ? (
        <div style={stateBoxStyle}>
          <span>📋</span>
          <span>No projects found in PMO_Projects list.</span>
        </div>
      ) : (
        <div>
          <table style={tableStyle}>
            <thead>
              <tr>
                <th style={thStyle}>#</th>
                <th style={thStyle}>Project Name</th>
                <th style={thStyle}>Status</th>
                <th style={thStyle}>Project Manager</th>
              </tr>
            </thead>
            <tbody>
              {projectList.map(function(project, index) {
                const badge = getStatusBadge(project.Project_Status);
                const managerName = project.Project_Manager && project.Project_Manager.Title
                  ? project.Project_Manager.Title
                  : '—';
                const initials = managerName !== '—' ? toInitials(managerName) : '?';
                const rowBg = index % 2 === 0 ? '#ffffff' : '#fafafa';

                const badgeStyle: React.CSSProperties = {
                  display: 'inline-block',
                  padding: '3px 10px',
                  borderRadius: '12px',
                  fontSize: '12px',
                  fontWeight: 600,
                  background: badge.background,
                  color: badge.color,
                  border: badge.border,
                };

                const avatarStyle: React.CSSProperties = {
                  width: '28px',
                  height: '28px',
                  borderRadius: '50%',
                  background: '#4f46e5',
                  color: '#fff',
                  display: 'inline-flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  fontSize: '11px',
                  fontWeight: 700,
                  marginRight: '8px',
                  verticalAlign: 'middle',
                };

                return (
                  <tr key={project.Id} style={{ background: rowBg }}>
                    <td style={{ ...tdStyle, color: '#9ca3af', width: '40px' }}>{index + 1}</td>
                    <td style={{ ...tdStyle, fontWeight: 500, color: '#111827' }}>{project.Title}</td>
                    <td style={tdStyle}>
                      <span style={badgeStyle}>{project.Project_Status}</span>
                    </td>
                    <td style={tdStyle}>
                      <span style={avatarStyle}>{initials}</span>
                      {managerName}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          {updatedLabel && (
            <p style={{ margin: '12px 0 0', fontSize: '11px', color: '#9ca3af', textAlign: 'right' }}>
              Last refreshed at {updatedLabel}
            </p>
          )}
        </div>
      )}
    </div>
  );
};

export default DashboardView;
