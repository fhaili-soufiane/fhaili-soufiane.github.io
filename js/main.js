document.addEventListener('DOMContentLoaded', () => {

  /* ── Mobile menu ── */
  const toggle = document.querySelector('.menu-toggle');
  const nav = document.querySelector('nav');
  if (toggle && nav) {
    toggle.addEventListener('click', () => nav.classList.toggle('open'));
    document.addEventListener('click', (e) => {
      if (!e.target.closest('header')) nav.classList.remove('open');
    });
  }

  /* ── Active nav link ── */
  const current = location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('nav a').forEach(a => {
    const href = a.getAttribute('href');
    if (href === current || (current === '' && href === 'index.html')) {
      a.classList.add('active');
    }
  });

  /* ── Scroll: navbar shadow ── */
  const header = document.querySelector('header');
  let lastScroll = 0;
  window.addEventListener('scroll', () => {
    const y = window.scrollY;
    if (y > 20) {
      header.classList.add('scrolled');
    } else {
      header.classList.remove('scrolled');
    }
    lastScroll = y;
  }, { passive: true });

  /* ── Fade-in on scroll (Intersection Observer) ── */
  const fadeEls = document.querySelectorAll('.fade-in');
  if (fadeEls.length && 'IntersectionObserver' in window) {
    const observer = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          entry.target.classList.add('visible');
          observer.unobserve(entry.target);
        }
      });
    }, { threshold: 0.1, rootMargin: '0px 0px -40px 0px' });

    fadeEls.forEach(el => observer.observe(el));
  } else {
    // fallback
    fadeEls.forEach(el => el.classList.add('visible'));
  }

});
