// Interactive behaviors for Lab 2 - Konstruct
document.addEventListener('DOMContentLoaded', () => {
  // 1. Smooth scroll for all internal links
  const links = document.querySelectorAll('a[href^="#"]');
  links.forEach(link => {
    link.addEventListener('click', e => {
      const targetId = link.getAttribute('href');
      if (targetId && targetId !== '#') {
        const targetElement = document.querySelector(targetId);
        if (targetElement) {
          e.preventDefault();
          targetElement.scrollIntoView({ behavior: 'smooth' });
        }
      }
    });
  });

  // 2. Active nav link on scroll
  const sections = document.querySelectorAll('header, section, footer');
  const navLinks = document.querySelectorAll('.navbar__link');

  window.addEventListener('scroll', () => {
    let current = '';
    const scrollPos = window.pageYOffset || document.documentElement.scrollTop;

    sections.forEach(section => {
      const sectionTop = section.offsetTop - 100;
      const sectionHeight = section.offsetHeight;
      if (scrollPos >= sectionTop && scrollPos < sectionTop + sectionHeight) {
        current = section.getAttribute('id');
      }
    });

    navLinks.forEach(link => {
      link.classList.remove('navbar__link--active');
      if (link.getAttribute('href') === `#${current}`) {
        link.classList.add('navbar__link--active');
      }
    });
  });

  // 3. Newsletter form handler
  const form = document.querySelector('.newsletter__form');
  if (form) {
    form.addEventListener('submit', e => {
      e.preventDefault();
      const input = form.querySelector('.newsletter__input');
      if (input && input.value) {
        alert(`Thank you for subscribing with ${input.value}!`);
        input.value = '';
      }
    });
  }
});
