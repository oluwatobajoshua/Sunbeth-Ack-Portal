import React, { useMemo, useRef } from 'react';

export type TabConfig = { id: string; label: string; icon?: string };

interface TabNavProps {
  tabs: TabConfig[];
  activeId: string;
  onChange: (id: string) => void;
}

const TabNav: React.FC<TabNavProps> = ({ tabs, activeId, onChange }) => {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const ids = useMemo(() => tabs.map(t => t.id), [tabs]);
  const activeIdx = Math.max(0, ids.indexOf(activeId));

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
      e.preventDefault();
      const dir = e.key === 'ArrowRight' ? 1 : -1;
      const next = (activeIdx + dir + ids.length) % ids.length;
      onChange(ids[next]);
      // Move focus to the newly active tab
      const el = containerRef.current?.querySelector(`[data-tab-id="${ids[next]}"]`) as HTMLButtonElement | null;
      el?.focus();
    }
  };

  return (
    <div ref={containerRef} role="tablist" aria-label="Admin sections" tabIndex={0} style={{ display: 'flex', gap: 4, marginBottom: 24, borderBottom: '2px solid #f0f0f0' }} onKeyDown={onKeyDown}>
      {tabs.map(t => {
        const isActive = t.id === activeId;
        return (
          <button
            key={t.id}
            data-tab-id={t.id}
            role="tab"
            aria-selected={isActive}
            aria-controls={`panel-${t.id}`}
            tabIndex={isActive ? 0 : -1}
            className={isActive ? 'btn sm' : 'btn ghost sm'}
            onClick={() => onChange(t.id)}
            style={{ borderRadius: '8px 8px 0 0' }}
          >
            {t.icon} {t.label}
          </button>
        );
      })}
    </div>
  );
};

export default TabNav;
