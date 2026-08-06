/* ============================================================
   YOGGIES — PRODUCTS PAGE JS
   ============================================================ */

/* Add-to-cart feedback */
document.querySelectorAll('.pg-cart-btn').forEach(btn => {
  btn.addEventListener('click', e => {
    e.stopPropagation();

    if (btn.classList.contains('pg-added')) return;

    btn.classList.add('pg-added');
    btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"
      stroke-linecap="round" stroke-linejoin="round" width="17" height="17">
      <polyline points="20 6 9 17 4 12"/>
    </svg>`;

    setTimeout(() => {
      btn.classList.remove('pg-added');
      btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"
        stroke-linecap="round" stroke-linejoin="round">
        <path d="M6 2L3 6v14a2 2 0 002 2h14a2 2 0 002-2V6l-3-4z"/>
        <line x1="3" y1="6" x2="21" y2="6"/>
        <path d="M16 10a4 4 0 01-8 0"/>
      </svg>`;
    }, 1400);
  });
});

/* Pagination active state */
document.querySelectorAll('.pg-page-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.pg-page-btn').forEach(b => b.classList.remove('pg-page-btn--active'));
    if (!['‹', '›'].includes(btn.textContent.trim())) {
      btn.classList.add('pg-page-btn--active');
    }
  });
});
