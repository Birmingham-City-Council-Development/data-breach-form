import * as React from 'react';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
} from 'recharts';
import { IBreachListRow } from '../../dataBreachForm/services/SharePointService';

// Recharts 2.x return types conflict with React 17's JSX checker — cast to silence.
/* eslint-disable @typescript-eslint/no-explicit-any */
const RC  = ResponsiveContainer as unknown as React.ComponentType<any>;
const BC  = BarChart           as unknown as React.ComponentType<any>;
const B   = Bar                as unknown as React.ComponentType<any>;
const XA  = XAxis              as unknown as React.ComponentType<any>;
const YA  = YAxis              as unknown as React.ComponentType<any>;
const CG  = CartesianGrid      as unknown as React.ComponentType<any>;
const TT  = Tooltip            as unknown as React.ComponentType<any>;
/* eslint-enable @typescript-eslint/no-explicit-any */

interface IStatsPanelProps {
  items: IBreachListRow[];
}

interface IDirEntry { name: string; count: number; }

const StatsPanel: React.FC<IStatsPanelProps> = ({ items }) => {
  const now = new Date();

  const total        = items.length;
  const open         = items.filter(i => !i.Status || i.Status === 'Open').length;
  const pending      = items.filter(i => i.Status === 'Pending').length;
  const closed       = items.filter(i => i.Status === 'Closed').length;
  const reportedIco  = items.filter(i => i.Reported_x0020_to_x0020_ICO === 'Yes').length;
  const last30       = items.filter(i => (now.getTime() - new Date(i.Created).getTime()) <= 30 * 86400000).length;
  const last90       = items.filter(i => (now.getTime() - new Date(i.Created).getTime()) <= 90 * 86400000).length;

  // Submissions per month — last 12 months
  const monthData: { month: string; count: number }[] = [];
  for (let m = 11; m >= 0; m--) {
    const target = new Date(now.getFullYear(), now.getMonth() - m, 1);
    const label  = target.toLocaleDateString('en-GB', { month: 'short', year: '2-digit' });
    const count  = items.filter(i => {
      const c = new Date(i.Created);
      return c.getFullYear() === target.getFullYear() && c.getMonth() === target.getMonth();
    }).length;
    monthData.push({ month: label, count });
  }

  // Breaches by directorate — top 10
  const byDir: Record<string, number> = {};
  items.forEach(i => {
    const d = i.Directorate_x0020_of_x0020_repor || 'Not specified';
    byDir[d] = (byDir[d] || 0) + 1;
  });
  const dirData: IDirEntry[] = Object.keys(byDir)
    .map((name: string) => ({ name, count: byDir[name] }))
    .sort((a: IDirEntry, b: IDirEntry) => b.count - a.count)
    .slice(0, 10);

  const card = (value: number, label: string, accent?: string): React.ReactElement => (
    <div style={{
      background: '#f3f2f1',
      border: '1px solid #b1b4b6',
      borderTop: accent ? `4px solid ${accent}` : '1px solid #b1b4b6',
      borderRadius: '4px',
      padding: '16px 20px',
      textAlign: 'center',
      minWidth: '110px',
    }}>
      <div style={{ fontSize: '2rem', fontWeight: 700, lineHeight: 1.1, color: accent || '#0b0c0c' }}>{value}</div>
      <div className="govuk-body-s" style={{ color: '#505a5f', marginBottom: 0 }}>{label}</div>
    </div>
  );

  return (
    <div style={{ marginBottom: '30px', borderBottom: '1px solid #b1b4b6', paddingBottom: '30px' }}>
      <h2 className="govuk-heading-m" style={{ marginBottom: '16px' }}>Statistics</h2>

      <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap', marginBottom: '30px' }}>
        {card(total,       'Total')}
        {card(open,        'Open',            '#d4351c')}
        {card(pending,     'Pending',         '#f47738')}
        {card(closed,      'Closed',          '#00703c')}
        {card(reportedIco, 'Reported to ICO', '#1d70b8')}
        {card(last30,      'Last 30 days')}
        {card(last90,      'Last 90 days')}
      </div>

      <div style={{ display: 'flex', gap: '40px', flexWrap: 'wrap' }}>
        <div style={{ flex: '1 1 340px', minWidth: 0 }}>
          <h3 className="govuk-heading-s">Submissions per month (last 12 months)</h3>
          <RC width="100%" height={200}>
            <BC data={monthData} margin={{ top: 4, right: 8, left: -20, bottom: 0 }}>
              <CG strokeDasharray="3 3" vertical={false} />
              <XA dataKey="month" tick={{ fontSize: 11 }} />
              <YA allowDecimals={false} tick={{ fontSize: 11 }} />
              <TT />
              <B dataKey="count" fill="#1d70b8" name="Breaches" radius={[2, 2, 0, 0]} />
            </BC>
          </RC>
        </div>

        {dirData.length > 0 && (
          <div style={{ flex: '1 1 340px', minWidth: 0 }}>
            <h3 className="govuk-heading-s">Breaches by directorate</h3>
            <RC width="100%" height={Math.max(200, dirData.length * 28)}>
              <BC data={dirData} layout="vertical" margin={{ top: 4, right: 20, left: 4, bottom: 0 }}>
                <CG strokeDasharray="3 3" horizontal={false} />
                <XA type="number" allowDecimals={false} tick={{ fontSize: 11 }} />
                <YA type="category" dataKey="name" tick={{ fontSize: 10 }} width={160} />
                <TT />
                <B dataKey="count" fill="#1d70b8" name="Breaches" radius={[0, 2, 2, 0]} />
              </BC>
            </RC>
          </div>
        )}
      </div>
    </div>
  );
};

export default StatsPanel;
