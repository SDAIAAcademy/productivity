/* AI Productivity — lab guide pages: logo, copy buttons on prompts, course navigation. */
(function () {
  const logo = document.getElementById('logo');
  if (logo) logo.innerHTML = window.ACADEMY_LOGO || 'SDAIA Academy';

  /* every .prompt gets a copy button; the copied text leaves out the tool label and the button itself */
  document.querySelectorAll('.prompt').forEach((p) => {
    const b = document.createElement('button');
    b.className = 'copy';
    b.type = 'button';
    b.textContent = 'نسخ';
    b.addEventListener('click', () => {
      const clone = p.cloneNode(true);
      clone.querySelectorAll('.copy').forEach((x) => x.remove());
      const text = clone.innerText.trim();
      const done = () => { b.textContent = 'تم النسخ ✓'; b.classList.add('ok'); setTimeout(() => { b.textContent = 'نسخ'; b.classList.remove('ok'); }, 1600); };
      if (navigator.clipboard) navigator.clipboard.writeText(text).then(done, () => fallback(text, done));
      else fallback(text, done);
    });
    p.appendChild(b);
  });
  function fallback(text, done) {
    const t = document.createElement('textarea');
    t.value = text;
    document.body.appendChild(t);
    t.select();
    try { document.execCommand('copy'); done(); } catch (e) { /* clipboard unavailable */ }
    t.remove();
  }
})();
