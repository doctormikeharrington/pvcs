// PVCS Internal Team Site — Main JavaScript

document.addEventListener('DOMContentLoaded', function () {
  // Mobile navigation toggle
  const navToggle = document.querySelector('.nav-toggle');
  const nav = document.querySelector('.nav');

  if (navToggle && nav) {
    navToggle.addEventListener('click', function () {
      nav.classList.toggle('nav--open');
      navToggle.classList.toggle('nav-toggle--active');
    });

    nav.querySelectorAll('.nav__link, .nav__menu-link').forEach(function (link) {
      link.addEventListener('click', function () {
        nav.classList.remove('nav--open');
        navToggle.classList.remove('nav-toggle--active');
      });
    });
  }

  // Highlight the current page in the nav
  const currentPage = window.location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.nav__link, .nav__menu-link').forEach(function (link) {
    const href = link.getAttribute('href');
    if (href === currentPage || (currentPage === '' && href === 'index.html')) {
      link.classList.add(link.classList.contains('nav__menu-link') ? 'is-active' : 'nav__link--active');
    }
  });

  // Keep the "Schedule" parent highlighted on either schedule page
  if (currentPage === 'schedule.html' || currentPage === 'schedule-pa.html') {
    const toggle = document.querySelector('.nav__group-toggle');
    if (toggle) toggle.classList.add('nav__link--active');
  }

  // Auto-open an accordion <details> when linked to directly (e.g. #connection-dialer)
  function openTargetedDetails() {
    const id = window.location.hash.slice(1);
    if (!id) return;
    const el = document.getElementById(id);
    if (el && el.tagName.toLowerCase() === 'details') {
      el.open = true;
      el.scrollIntoView();
    }
  }
  openTargetedDetails();
  window.addEventListener('hashchange', openTargetedDetails);
});
