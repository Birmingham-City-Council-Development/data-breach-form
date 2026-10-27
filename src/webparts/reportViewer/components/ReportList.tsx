import * as React from 'react';
import { IBreachListRow } from '../../dataBreachForm/services/SharePointService';
import { SortField } from './reportViewerTypes';
import { formatDate, getRef, STATUS_ORDER, LIST_PAGE_SIZE } from './reportViewerUtils';
import { downloadCsv } from './reportExport';
import StatsPanel from './StatsPanel';

type SearchBy = 'Data_x0020_Breach_x0020_ID' | 'Author';

export interface IReportListControls {
  searchText: string;
  setSearchText: (value: string) => void;
  searchBy: SearchBy;
  setSearchBy: (value: SearchBy) => void;
  showStats: boolean;
  toggleStats: () => void;
  filteredItems: IBreachListRow[];
  visibleItems: IBreachListRow[];
  toggleSort: (field: SortField) => void;
  sortIcon: (field: SortField) => string;
  loadMore: () => void;
}

// Search/sort/paging state for the list. Called from ReportViewer rather than
// inside ReportList so it survives opening a report and coming back — the
// list view unmounts while a report is open.
export function useReportListControls(items: IBreachListRow[]): IReportListControls {
  const [sortField, setSortField] = React.useState<SortField>('Created');
  const [sortDir, setSortDir] = React.useState<'asc' | 'desc'>('desc');
  const [searchText, setSearchText] = React.useState('');
  const [searchBy, setSearchBy] = React.useState<SearchBy>('Data_x0020_Breach_x0020_ID');
  const [visibleCount, setVisibleCount] = React.useState(LIST_PAGE_SIZE);
  const [showStats, setShowStats] = React.useState(true);

  const toggleSort = (field: SortField): void => {
    if (field === sortField) {
      setSortDir(function(d) { return d === 'asc' ? 'desc' : 'asc'; });
    } else {
      setSortField(field);
      setSortDir('asc');
    }
  };

  const sortedItems = items.slice().sort(function(a, b) {
    let cmp = 0;
    // Data_x0020_Breach_x0020_ID no longer reflects a real value (the flow
    // that populated it is off) — the reference is just "DBN-" + Id now, so
    // sorting "by reference" is the same as sorting by Id numerically. The
    // sortField name is kept as-is; it's just the internal key for this
    // column, not a read of the SharePoint field anymore.
    if (sortField === 'Data_x0020_Breach_x0020_ID') {
      cmp = a.Id - b.Id;
    } else if (sortField === 'Status') {
      const oA = STATUS_ORDER[a.Status] !== undefined ? STATUS_ORDER[a.Status] : 99;
      const oB = STATUS_ORDER[b.Status] !== undefined ? STATUS_ORDER[b.Status] : 99;
      cmp = oA - oB;
    } else if (sortField === 'Created') {
      cmp = new Date(a.Created).getTime() - new Date(b.Created).getTime();
    } else if (sortField === 'Directorate_x0020_of_x0020_repor') {
      const dA = (a.Directorate_x0020_of_x0020_repor || '').toLowerCase();
      const dB = (b.Directorate_x0020_of_x0020_repor || '').toLowerCase();
      cmp = dA < dB ? -1 : dA > dB ? 1 : 0;
    }
    return sortDir === 'asc' ? cmp : -cmp;
  });

  const filteredItems = sortedItems.filter(function(item) {
    if (!searchText.trim()) return true;
    const q = searchText.trim().toLowerCase();
    if (searchBy === 'Data_x0020_Breach_x0020_ID') return getRef(item.Id).toLowerCase().indexOf(q) !== -1;
    if (searchBy === 'Author') return (item.Author ? item.Author.Title : '').toLowerCase().indexOf(q) !== -1;
    return true;
  });

  React.useEffect(function() {
    setVisibleCount(LIST_PAGE_SIZE);
  }, [searchText, searchBy, sortField, sortDir]);

  const sortIcon = (field: SortField): string => {
    if (field !== sortField) return ' ↕';
    return sortDir === 'asc' ? ' ▲' : ' ▼';
  };

  return {
    searchText,
    setSearchText,
    searchBy,
    setSearchBy,
    showStats,
    toggleStats: function() { setShowStats(function(s) { return !s; }); },
    filteredItems,
    visibleItems: filteredItems.slice(0, visibleCount),
    toggleSort,
    sortIcon,
    loadMore: function() { setVisibleCount(function(c) { return c + LIST_PAGE_SIZE; }); },
  };
}

const ReportList: React.FC<{
  items: IBreachListRow[];
  loading: boolean;
  isAdmin: boolean;
  controls: IReportListControls;
  onOpen: (id: number) => void;
}> = ({ items, loading, isAdmin, controls, onOpen }) => {
  const { searchText, setSearchText, searchBy, setSearchBy, showStats, toggleStats, filteredItems, visibleItems, toggleSort, sortIcon, loadMore } = controls;
  return (
    <div style={{ padding: '20px' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', marginBottom: '10px' }}>
        <h1 className="govuk-heading-l" style={{ marginBottom: 0 }}>Data Breach Reports</h1>
        <div style={{ display: 'flex', gap: '10px' }}>
          {isAdmin && (
            <button
              type="button"
              className="govuk-button govuk-button--secondary"
              style={{ marginBottom: 0 }}
              onClick={toggleStats}
            >
              {showStats ? 'Hide statistics' : 'Show statistics'}
            </button>
          )}
          {!loading && items.length > 0 && (
            <button
              type="button"
              className="govuk-button govuk-button--secondary"
              style={{ marginBottom: 0 }}
              onClick={function() { downloadCsv(filteredItems); }}
            >
              Export CSV
            </button>
          )}
        </div>
      </div>

      {isAdmin && (
        <div className="govuk-inset-text" style={{ marginBottom: '20px' }}>
          You are viewing as an administrator and can see all submissions.
        </div>
      )}

      {isAdmin && showStats && !loading && items.length > 0 && (
        <StatsPanel items={items} />
      )}

      <div style={{ display: 'flex', gap: '10px', alignItems: 'flex-end', marginBottom: '20px' }}>
        <div className="govuk-form-group" style={{ marginBottom: 0, flexShrink: 0 }}>
          <label className="govuk-label govuk-body-s" style={{ marginBottom: '4px' }}>Search by</label>
          <select
            className="govuk-select"
            value={searchBy}
            onChange={function(e) { setSearchBy(e.target.value as SearchBy); setSearchText(''); }}
          >
            <option value="Data_x0020_Breach_x0020_ID">Reference</option>
            <option value="Author">Submitted by</option>
          </select>
        </div>
        <div className="govuk-form-group" style={{ marginBottom: 0, flexGrow: 1 }}>
          <label className="govuk-label govuk-body-s" style={{ marginBottom: '4px' }}>Search</label>
          <input
            className="govuk-input"
            type="text"
            value={searchText}
            placeholder={searchBy === 'Data_x0020_Breach_x0020_ID' ? 'e.g. DBN-12' : 'Type to filter…'}
            onChange={function(e) { setSearchText(e.target.value); }}
          />
        </div>
        {searchText && (
          <button
            type="button"
            className="govuk-button govuk-button--secondary"
            style={{ marginBottom: 0 }}
            onClick={function() { setSearchText(''); }}
          >
            Clear
          </button>
        )}
      </div>
      {loading ? (
        <p className="govuk-body">Loading…</p>
      ) : items.length === 0 ? (
        <p className="govuk-body">No breach reports found.</p>
      ) : filteredItems.length === 0 ? (
        <p className="govuk-body">No results match your search.</p>
      ) : (
        <table className="govuk-table">
          <thead className="govuk-table__head">
            <tr className="govuk-table__row">
              <th className="govuk-table__header" style={{ cursor: 'pointer', whiteSpace: 'nowrap' }} onClick={function() { toggleSort('Data_x0020_Breach_x0020_ID'); }}>
                Reference{sortIcon('Data_x0020_Breach_x0020_ID')}
              </th>
              <th className="govuk-table__header" style={{ cursor: 'pointer', whiteSpace: 'nowrap' }} onClick={function() { toggleSort('Status'); }}>
                Status{sortIcon('Status')}
              </th>
              <th className="govuk-table__header">Submitted by</th>
              <th className="govuk-table__header" style={{ cursor: 'pointer', whiteSpace: 'nowrap' }} onClick={function() { toggleSort('Created'); }}>
                Date{sortIcon('Created')}
              </th>
              <th className="govuk-table__header" style={{ cursor: 'pointer', whiteSpace: 'nowrap' }} onClick={function() { toggleSort('Directorate_x0020_of_x0020_repor'); }}>
                Directorate{sortIcon('Directorate_x0020_of_x0020_repor')}
              </th>
              <th className="govuk-table__header" />
            </tr>
          </thead>
          <tbody className="govuk-table__body">
            {visibleItems.map(function(item) {
              return (
                <tr key={item.Id} className="govuk-table__row">
                  <td className="govuk-table__cell">
                    <strong>{getRef(item.Id)}</strong>
                  </td>
                  <td className="govuk-table__cell">
                    <strong className={`govuk-tag${item.Status === 'Closed' ? ' govuk-tag--green' : item.Status === 'Pending' ? ' govuk-tag--yellow' : ''}`}>
                      {item.Status || 'Open'}
                    </strong>
                  </td>
                  <td className="govuk-table__cell">{item.Author ? item.Author.Title : ''}</td>
                  <td className="govuk-table__cell">{formatDate(item.Created)}</td>
                  <td className="govuk-table__cell">{item.Directorate_x0020_of_x0020_repor || ''}</td>
                  <td className="govuk-table__cell">
                    <a className="govuk-link" href="#" onClick={function(e) { e.preventDefault(); onOpen(item.Id); }}>
                      View
                    </a>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}
      {!loading && filteredItems.length > visibleItems.length && (
        <div style={{ textAlign: 'center', marginTop: '16px' }}>
          <p className="govuk-body-s" style={{ color: '#505a5f' }}>
            Showing {visibleItems.length} of {filteredItems.length}
          </p>
          <button
            type="button"
            className="govuk-button govuk-button--secondary"
            onClick={loadMore}
          >
            Load more
          </button>
        </div>
      )}
    </div>
  );
};

export default ReportList;
