import * as React from 'react';

import { SharePointListService, IProject } from '../services/SharePointListService';
import { useQuery, IDashboardViewProps } from '../hooks/useQuery';

// ─── Status Badge ─────────────────────────────────────────────────────────────

var STATUS_COLORS: { [key: string]: { bg: string; color: string; border: string } } = {
  'Active':    { bg: '#d1fae5', color: '#065f46', border: '1px solid #6ee7b7' },
  'On Hold':   { bg: '#fef3c7', color: '#92400e', border: '1px solid #fcd34d' },
  'Completed': { bg: '#dbeafe', color: '#1e3a8a', border: '1px solid #93c5fd' },
  'Cancelled': { bg: '#fee2e2', color: '#991b1b', border: '1px solid #fca5a5' }
};

function getStatusStyle(status: string): { bg: string; color: string; border: string } {
  return STATUS_COLORS[status] || { bg: '#f3f4f6', color: '#374151', border: '1px solid #d1d5db' };
}

function toInitials(name: string): string {
  var parts = name.split(' ').filter(function(w) { return w.length > 0; });
  return parts.slice(0, 2).map(function(w) { return w[0].toUpperCase(); }).join('');
}

// ─── Component ────────────────────────────────────────────────────────────────

var DashboardView: React.FC<IDashboardViewProps> = function(props) {
  var context = props.context;
  var title = props.title;

  var service = React.useMemo(function() {
    return new SharePointListService(context);
  }, [context]);

  var result = useQuery<IProject[], Error>(function() {
    return service.getProjects();
  });

  var projects = result.data;
  var isLoading = result.isLoading;
  var isError = result.isError;
  var error = result.error;
  var dataUpdatedAt = result.dataUpdatedAt;

  // Inject spinner keyframes once
  React.useEffect(function() {
    var id = 'pmo-spin-style';
    if (!document.getElementById(id)) {
      var el = document.createElement('style');
      el.id = id;
      el.textContent = '@keyframes pmo-spin { to { transform: rotate(360deg); } }';
      document.head.appendChild(el);
    }
  }, []);

  var wrapStyle: React.CSSProperties = {
    fontFamily: "'Segoe UI', -apple-system, sans-serif",
    padding: '24px',
    background: '#fff',
    minHeight: '200px'
  };

  if (isLoading) {
    return React.createElement('div', { style: wrapStyle },
      React.createElement('div', {
        style: {
          display: 'flex', flexDirection: 'column', alignItems: 'center',
          justifyContent: 'center', gap: '8px', padding: '48px 0', color: '#6b7280'
        }
      },
        React.createElement('div', {
          style: {
            width: '36px', height: '36px', borderRadius: '50%',
            border: '3px solid #e5e7eb', borderTop: '3px solid #4f46e5',
            animation: 'pmo-spin 0.8s linear infinite'
          }
        }),
        React.createElement('span', null, 'Loading projects...')
      )
    );
  }

  if (isError) {
    return React.createElement('div', { style: wrapStyle },
      React.createElement('div', {
        style: {
          background: '#fff1f2', border: '1px solid #fecdd3',
          borderRadius: '8px', padding: '16px 20px', color: '#9f1239', fontSize: '14px'
        }
      },
        React.createElement('strong', null, 'Failed to load PMO Projects'),
        React.createElement('br'),
        error ? error.message : 'An unknown error occurred.'
      )
    );
  }

  var projectList: IProject[] = projects || [];
  var updatedLabel = dataUpdatedAt ? new Date(dataUpdatedAt).toLocaleTimeString() : '';

  var thStyle: React.CSSProperties = {
    textAlign: 'left', padding: '10px 14px', background: '#f9fafb',
    color: '#6b7280', fontWeight: 600, fontSize: '12px',
    textTransform: 'uppercase', letterSpacing: '0.05em', borderBottom: '2px solid #e5e7eb'
  };

  var tdStyle: React.CSSProperties = {
    padding: '12px 14px', borderBottom: '1px solid #f3f4f6',
    color: '#374151', verticalAlign: 'middle'
  };

  return React.createElement('div', { style: wrapStyle },

    // Header
    React.createElement('div', {
      style: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '20px' }
    },
      React.createElement('h2', { style: { fontSize: '20px', fontWeight: 600, color: '#111827', margin: 0 } }, title),
      projectList.length > 0 && React.createElement('span', {
        style: {
          fontSize: '12px', fontWeight: 500, background: '#eff6ff',
          color: '#1d4ed8', border: '1px solid #bfdbfe', borderRadius: '12px', padding: '2px 10px'
        }
      }, projectList.length + ' project' + (projectList.length !== 1 ? 's' : ''))
    ),

    // Empty state
    projectList.length === 0
      ? React.createElement('div', {
          style: { display: 'flex', flexDirection: 'column', alignItems: 'center', padding: '48px 0', color: '#6b7280' }
        },
          React.createElement('span', null, 'No projects found in PMO_Projects list.')
        )

      // Table
      : React.createElement('div', null,
          React.createElement('table', { style: { width: '100%', borderCollapse: 'collapse', fontSize: '14px' } },
            React.createElement('thead', null,
              React.createElement('tr', null,
                React.createElement('th', { style: thStyle }, '#'),
                React.createElement('th', { style: thStyle }, 'Project Name'),
                React.createElement('th', { style: thStyle }, 'Status'),
                React.createElement('th', { style: thStyle }, 'Project Manager')
              )
            ),
            React.createElement('tbody', null,
              projectList.map(function(project: IProject, index: number) {
                var sc = getStatusStyle(project.Project_Status);
                var managerName = (project.Project_Manager && project.Project_Manager.Title)
                  ? project.Project_Manager.Title : '—';
                var initials = managerName !== '—' ? toInitials(managerName) : '?';
                var rowBg = index % 2 === 0 ? '#fff' : '#fafafa';

                return React.createElement('tr', { key: project.Id, style: { background: rowBg } },
                  React.createElement('td', { style: Object.assign({}, tdStyle, { color: '#9ca3af', width: '40px' }) }, index + 1),
                  React.createElement('td', { style: Object.assign({}, tdStyle, { fontWeight: 500, color: '#111827' }) }, project.Title),
                  React.createElement('td', { style: tdStyle },
                    React.createElement('span', {
                      style: {
                        display: 'inline-block', padding: '3px 10px', borderRadius: '12px',
                        fontSize: '12px', fontWeight: 600,
                        background: sc.bg, color: sc.color, border: sc.border
                      }
                    }, project.Project_Status)
                  ),
                  React.createElement('td', { style: tdStyle },
                    React.createElement('span', {
                      style: {
                        display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                        width: '28px', height: '28px', borderRadius: '50%',
                        background: '#4f46e5', color: '#fff',
                        fontSize: '11px', fontWeight: 700, marginRight: '8px', verticalAlign: 'middle'
                      }
                    }, initials),
                    managerName
                  )
                );
              })
            )
          ),
          updatedLabel && React.createElement('p', {
            style: { margin: '12px 0 0', fontSize: '11px', color: '#9ca3af', textAlign: 'right' }
          }, 'Last refreshed at ' + updatedLabel)
        )
  );
};

export default DashboardView;
