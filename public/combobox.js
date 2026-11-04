/**
 * Combobox sederhana: ketik untuk mencari (dengan saran) atau klik panah untuk daftar.
 * getOptions(text, browse) -> [{ label, sub, data }]
 */
export function createCombo({ input, toggle, list, getOptions, onPick, onCommit }) {
  let items = [], active = -1, isOpen = false;

  function render() {
    list.replaceChildren();
    items.forEach((item, index) => {
      const button = document.createElement('button');
      button.type = 'button'; button.className = 'opt'; button.setAttribute('role', 'option');
      if (index === active) button.classList.add('active');
      button.textContent = item.label;
      if (item.sub) { const small = document.createElement('small'); small.textContent = item.sub; button.append(small); }
      button.addEventListener('mousedown', event => event.preventDefault());
      button.addEventListener('click', () => choose(item));
      list.append(button);
    });
    if (!items.length) { const empty = document.createElement('div'); empty.className = 'empty'; empty.textContent = 'Tidak ada pilihan.'; list.append(empty); }
    list.hidden = !isOpen;
    input.setAttribute('aria-expanded', String(isOpen));
    if (toggle) toggle.setAttribute('aria-expanded', String(isOpen));
  }
  function open(browse) { items = getOptions(input.value, browse); active = -1; isOpen = true; render(); }
  function close() { isOpen = false; render(); }
  function choose(item) { close(); onPick(item.data); }

  input.addEventListener('input', () => open(false));
  input.addEventListener('focus', () => input.select());
  input.addEventListener('blur', () => { setTimeout(() => { if (isOpen) close(); onCommit?.(input.value); }, 120); });
  input.addEventListener('keydown', event => {
    if (event.key === 'ArrowDown') { event.preventDefault(); if (!isOpen) open(true); else { active = Math.min(active + 1, items.length - 1); render(); list.querySelector('.active')?.scrollIntoView({ block: 'nearest' }); } }
    else if (event.key === 'ArrowUp') { event.preventDefault(); active = Math.max(active - 1, 0); render(); list.querySelector('.active')?.scrollIntoView({ block: 'nearest' }); }
    else if (event.key === 'Enter') { event.preventDefault(); if (isOpen && items.length) choose(items[Math.max(active, 0)]); else { close(); onCommit?.(input.value); } }
    else if (event.key === 'Escape') { if (isOpen) { event.stopPropagation(); close(); } }
  });
  toggle?.addEventListener('mousedown', event => event.preventDefault());
  toggle?.addEventListener('click', () => { if (isOpen) close(); else { input.focus(); open(true); } });

  return { close, setText(value) { input.value = value ?? ''; } };
}
