(() => {
  const root = document.querySelector('body.special-light-v1[data-special-innovation="v4"]');
  if (!root) return;

  const productWindow = document.querySelector('.product-window');
  if (productWindow && !productWindow.querySelector('.iv-sheet-tabs')) {
    const sheetTabs = document.createElement('div');
    sheetTabs.className = 'iv-sheet-tabs';
    sheetTabs.setAttribute('aria-label', 'Temsili çalışma sayfaları');
    sheetTabs.innerHTML = [
      ['NAKİT', 'green'],
      ['CARİ', 'blue'],
      ['BANKA', 'amber'],
      ['VALÖR', 'coral'],
    ].map(([label, tone]) => `<span class="iv-sheet-tab" data-tone="${tone}">${label}</span>`).join('');
    productWindow.appendChild(sheetTabs);
  }

  const reduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  const revealTargets = [
    ...document.querySelectorAll('.solution-card, .proof-stat, .benefit, .process-card, .why-card, .compare-shell, .cta-shell, .iv-lab-shell'),
  ];
  revealTargets.forEach((el) => el.classList.add('iv-reveal'));

  if (reduced || !('IntersectionObserver' in window)) {
    revealTargets.forEach((el) => el.classList.add('is-visible'));
  } else {
    const observer = new IntersectionObserver((entries, obs) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;
        entry.target.classList.add('is-visible');
        obs.unobserve(entry.target);
      });
    }, { threshold: 0.12, rootMargin: '0px 0px -7% 0px' });
    revealTargets.forEach((el) => observer.observe(el));
  }

  const tabs = [...document.querySelectorAll('.iv-lab-tab')];
  const panels = [...document.querySelectorAll('.iv-lab-panel')];
  const activate = (targetId, focus = false) => {
    tabs.forEach((tab) => {
      const active = tab.getAttribute('aria-controls') === targetId;
      tab.setAttribute('aria-selected', active ? 'true' : 'false');
      tab.tabIndex = active ? 0 : -1;
      if (active && focus) tab.focus();
    });
    panels.forEach((panel) => {
      const active = panel.id === targetId;
      panel.classList.toggle('is-active', active);
      panel.hidden = !active;
    });
  };

  tabs.forEach((tab, index) => {
    tab.addEventListener('click', () => activate(tab.getAttribute('aria-controls')));
    tab.addEventListener('keydown', (event) => {
      if (!['ArrowDown', 'ArrowRight', 'ArrowUp', 'ArrowLeft', 'Home', 'End'].includes(event.key)) return;
      event.preventDefault();
      let next = index;
      if (event.key === 'ArrowDown' || event.key === 'ArrowRight') next = (index + 1) % tabs.length;
      if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') next = (index - 1 + tabs.length) % tabs.length;
      if (event.key === 'Home') next = 0;
      if (event.key === 'End') next = tabs.length - 1;
      activate(tabs[next].getAttribute('aria-controls'), true);
    });
  });

  const current = tabs.find((tab) => tab.getAttribute('aria-selected') === 'true') || tabs[0];
  if (current) activate(current.getAttribute('aria-controls'));
})();
