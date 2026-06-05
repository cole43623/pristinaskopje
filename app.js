// ═══════════════════════════════════════
//  NAV: mobile menu toggle + scrolled effect
// ═══════════════════════════════════════
const navbar     = document.getElementById('navbar');
const hamburger  = document.getElementById('hamburger');
const mobileMenu = document.getElementById('mobileMenu');

window.addEventListener('scroll', () => {
  navbar.classList.toggle('scrolled', window.scrollY > 50);
}, { passive: true });

hamburger.addEventListener('click', () => {
  hamburger.classList.toggle('open');
  mobileMenu.classList.toggle('open');
});

mobileMenu.querySelectorAll('a').forEach(link => {
  link.addEventListener('click', () => {
    hamburger.classList.remove('open');
    mobileMenu.classList.remove('open');
  });
});


// ═══════════════════════════════════════
//  INTERACTIVE TELA: Watercolor Mouse Trail
// ═══════════════════════════════════════
const canvasContainer = document.getElementById('paint-canvas-container');
const paintColors = [
  'rgba(255, 224, 230, 0.45)', // soft rose
  'rgba(208, 236, 252, 0.45)', // soft ice blue
  'rgba(253, 232, 206, 0.45)', // soft warm gold
  'rgba(213, 243, 219, 0.45)', // soft green
  'rgba(255, 249, 214, 0.45)'  // soft yellow
];

let lastDraw = 0;
const drawInterval = 28; // millisecondi tra una bolla e l'altra (flessibile e leggero)

function createWatercolorDrop(x, y, isClick = false) {
  if (!canvasContainer) return;

  const drop = document.createElement('div');
  drop.className = 'paint-drop';
  
  // Calcolo dimensioni casuali
  const size = isClick 
    ? Math.random() * 24 + 14 
    : Math.random() * 22 + 10;
  
  drop.style.width = `${size}px`;
  drop.style.height = `${size}px`;
  
  // Scelta del colore casuale dalla tavolozza
  const randomColor = paintColors[Math.floor(Math.random() * paintColors.length)];
  drop.style.background = randomColor;

  // Coordinate
  const offsetX = isClick ? (Math.random() * 40 - 20) : (Math.random() * 8 - 4);
  const offsetY = isClick ? (Math.random() * 40 - 20) : (Math.random() * 8 - 4);
  
  drop.style.left = `${x + offsetX}px`;
  drop.style.top = `${y + offsetY}px`;

  // Se è scaturito da un click, applichiamo un moto di espansione esplosivo personalizzato
  if (isClick) {
    drop.style.animationDuration = '1.3s';
    drop.style.filter = 'blur(3px)';
  }

  canvasContainer.appendChild(drop);

  // Rimuovi dal DOM al termine dell'animazione
  setTimeout(() => {
    drop.remove();
  }, isClick ? 1300 : 1100);
}

// Disegna al movimento del mouse (solo desktop per non interferire con il touch-scroll)
window.addEventListener('mousemove', (e) => {
  if (window.innerWidth <= 640) return;
  
  const now = Date.now();
  if (now - lastDraw > drawInterval) {
    createWatercolorDrop(e.clientX, e.clientY, false);
    lastDraw = now;
  }
});

// Esplosione a spruzzo sul click
window.addEventListener('click', (e) => {
  // Ignoriamo i click sugli elementi interattivi (pulsanti, link, card) per non coprire l'azione principale
  if (e.target.closest('a, button, .sticky-note, .paint-modal-content')) return;

  // Genera un cluster di bolle d'acqua
  const burstCount = Math.random() * 5 + 6;
  for (let i = 0; i < burstCount; i++) {
    createWatercolorDrop(e.clientX, e.clientY, true);
  }
});


// ═══════════════════════════════════════
//  PAINT MODAL: Popup che germoglia
// ═══════════════════════════════════════
const paintModal = document.getElementById('paintModal');
const modalClose = document.getElementById('modalClose');
const modalBody  = document.getElementById('modalBody');
const modalContent = document.querySelector('.paint-modal-content');

// Raccoglie tutti gli elementi cliccabili
const actCards = document.querySelectorAll('.act-card');

actCards.forEach(card => {
  card.addEventListener('click', () => {
    const title = card.getAttribute('data-title');
    const cost  = card.getAttribute('data-cost');
    const time  = card.getAttribute('data-time');
    const text  = card.getAttribute('data-text');
    const tips  = card.getAttribute('data-tips');
    const img   = card.getAttribute('data-image');

    // Blocco foto (solo se data-image è presente)
    const photoHTML = img ? `
      <div class="modal-photo-frame">
        <img
          src="${img}"
          alt="${title}"
          class="modal-photo"
          onerror="this.closest('.modal-photo-frame').style.display='none'"
        />
        <span class="modal-photo-caption">📸 ${title}</span>
      </div>
    ` : '';

    // Costruisce la bolla informativa
    modalBody.innerHTML = `
      ${photoHTML}
      <span class="modal-kicker">Consigli di viaggio · ${time} 💡</span>
      <h3 class="modal-title">${title}</h3>
      <div class="modal-meta-row">
        <span>🕒 ${time}</span>
        <span>💵 Costi: ${cost}</span>
      </div>
      <p class="modal-desc">${text}</p>
      ${tips ? `<div class="modal-tips">${tips}</div>` : ''}
    `;

    // Apre il popup con effetto germoglio
    paintModal.classList.add('open');
  });
});

// Chiusura del modale
function closeModal() {
  paintModal.classList.remove('open');
}

if (modalClose) {
  modalClose.addEventListener('click', closeModal);
}

// Chiude se si clicca fuori dalla nuvoletta di vernice
paintModal.addEventListener('click', (e) => {
  if (e.target === paintModal) {
    closeModal();
  }
});

// Chiude premendo ESC
window.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && paintModal.classList.contains('open')) {
    closeModal();
  }
});

// Swipe-to-close on mobile: chiusura trascinando lateralmente la modale
if (modalContent && paintModal) {
  let touchStartX = 0;
  let touchStartY = 0;
  let touchStartTime = 0;
  let currentTranslateX = 0;
  let isTouching = false;

  function onModalTouchStart(e) {
    if (!paintModal.classList.contains('open')) return;
    if (window.innerWidth > 640) return; // solo mobile
    const t = e.touches[0];
    touchStartX = t.clientX;
    touchStartY = t.clientY;
    touchStartTime = Date.now();
    currentTranslateX = 0;
    isTouching = true;
    modalContent.style.transition = 'none';
  }

  function onModalTouchMove(e) {
    if (!isTouching) return;
    const t = e.touches[0];
    const dx = t.clientX - touchStartX;
    const dy = t.clientY - touchStartY;
    // Se prevale lo scorrimento verticale, annulliamo la gesture orizzontale
    if (Math.abs(dy) > Math.abs(dx)) {
      isTouching = false;
      modalContent.style.transition = '';
      modalContent.style.transform = '';
      paintModal.style.background = '';
      return;
    }
    e.preventDefault();
    currentTranslateX = dx;
    const rotate = dx / 20;
    modalContent.style.transform = `translateX(${dx}px) translateZ(0) rotate(${rotate}deg)`;
    const progress = Math.min(Math.abs(dx) / (window.innerWidth * 0.6), 1);
    paintModal.style.background = `rgba(21, 48, 78, ${0.4 * (1 - progress)})`;
  }

  function onModalTouchEnd() {
    if (!isTouching) return;
    isTouching = false;
    const dx = currentTranslateX;
    const dt = Date.now() - touchStartTime;
    const velocity = dt > 0 ? dx / dt : 0;
    const absDx = Math.abs(dx);
    const threshold = 80; // px soglia per chiudere
    const velocityThreshold = 0.3; // px/ms
    modalContent.style.transition = 'transform 220ms cubic-bezier(0.2,0.8,0.2,1)';
    if (absDx > threshold || Math.abs(velocity) > velocityThreshold) {
      const dir = dx > 0 ? 1 : -1;
      const endX = dir * (window.innerWidth + 120);
      modalContent.style.transform = `translateX(${endX}px)`;
      paintModal.style.transition = 'opacity 220ms';
      paintModal.style.opacity = '0';
      setTimeout(() => {
        modalContent.style.transition = '';
        modalContent.style.transform = '';
        paintModal.style.transition = '';
        paintModal.style.opacity = '';
        paintModal.style.background = '';
        closeModal();
      }, 220);
    } else {
      // ritorna allo stato iniziale
      modalContent.style.transform = '';
      paintModal.style.background = '';
    }
  }

  paintModal.addEventListener('touchstart', onModalTouchStart, { passive: false });
  paintModal.addEventListener('touchmove', onModalTouchMove, { passive: false });
  paintModal.addEventListener('touchend', onModalTouchEnd);
  paintModal.addEventListener('touchcancel', onModalTouchEnd);
}

// Swipe-to-close on mobile for paint modal (touch gestures)
(() => {
  const modalContent = document.querySelector('.paint-modal-content');
  if (!modalContent) return;

  let startX = 0, startY = 0, startTime = 0, isTouching = false;
  const SWIPE_THRESHOLD = 80; // px
  const SWIPE_VELOCITY = 0.4; // px/ms

  function onTouchStart(e) {
    if (window.innerWidth > 640) return;
    if (!paintModal.classList.contains('open')) return;
    const t = e.touches ? e.touches[0] : e;
    startX = t.clientX;
    startY = t.clientY;
    startTime = Date.now();
    isTouching = true;
    modalContent.style.transition = 'none';
  }

  function onTouchMove(e) {
    if (!isTouching) return;
    const t = e.touches ? e.touches[0] : e;
    const dx = t.clientX - startX;
    const dy = t.clientY - startY;

    // if vertical movement is greater, don't treat as horizontal swipe
    if (Math.abs(dy) > Math.abs(dx)) return;

    // prevent page scroll while swiping horizontally
    e.preventDefault();

    modalContent.style.transform = `translateX(${dx}px) scale(1) rotate(0deg)`;
    const opacityFactor = 1 - Math.min(Math.abs(dx) / (window.innerWidth * 0.8), 1);
    paintModal.style.background = `rgba(21, 48, 78, ${0.4 * opacityFactor})`;
  }

  function onTouchEnd(e) {
    if (!isTouching) return;
    isTouching = false;
    const t = (e.changedTouches && e.changedTouches[0]) || e;
    const dx = t.clientX - startX;
    const dt = Date.now() - startTime;
    const velocity = Math.abs(dx) / (dt || 1);

    modalContent.style.transition = '';

    if (Math.abs(dx) > SWIPE_THRESHOLD || velocity > SWIPE_VELOCITY) {
      const dir = dx > 0 ? 1 : -1;
      modalContent.style.transform = `translateX(${dir * window.innerWidth}px) scale(0.9) rotate(${dir * 6}deg)`;
      paintModal.style.background = 'rgba(21, 48, 78, 0)';
      setTimeout(() => {
        closeModal();
        modalContent.style.transform = '';
        paintModal.style.background = '';
      }, 260);
    } else {
      modalContent.style.transform = '';
      paintModal.style.background = '';
    }
  }

  modalContent.addEventListener('touchstart', onTouchStart, { passive: true });
  modalContent.addEventListener('touchmove', onTouchMove, { passive: false });
  modalContent.addEventListener('touchend', onTouchEnd, { passive: true });
  modalContent.addEventListener('touchcancel', onTouchEnd, { passive: true });
})();


// ═══════════════════════════════════════
//  REVEAL ONDATA (Intersection Observer)
// ═══════════════════════════════════════
const revealEls = document.querySelectorAll(`
  .reveal, 
  .reveal-l, 
  .reveal-r, 
  .polaroid-wrapper, 
  .sticky-note, 
  .travel-ticket, 
  .travel-receipt-container, 
  .passport-notebook,
  .tips-notebook
`);

// Calcola un delay progressivo naturale per elementi adiacenti
function getOrganicDelay(el) {
  const p = el.parentElement;
  if (!p) return 0;

  // Trova tutti i fratelli rivelabili all'interno del genitore
  const siblings = Array.from(p.children).filter(child => {
    return child.classList.contains('reveal') ||
           child.classList.contains('reveal-l') ||
           child.classList.contains('reveal-r') ||
           child.classList.contains('polaroid-wrapper') ||
           child.classList.contains('sticky-note') ||
           child.classList.contains('travel-ticket') ||
           child.classList.contains('passport-notebook');
  });
  
  const idx = siblings.indexOf(el);
  if (idx < 0) return 0;

  // Staggering differenziato per blocco
  if (p.classList.contains('activities-canvas')) return idx * 110;
  if (p.classList.contains('rb-inner')) return idx * 100;
  if (p.tagName === 'UL') return idx * 65;

  return Math.min(idx * 80, 300);
}

const observer = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    if (!entry.isIntersecting) return;

    const delay = getOrganicDelay(entry.target);
    setTimeout(() => {
      entry.target.classList.add('is-visible');
    }, delay);

    // Smetti di osservare una volta rivelato
    observer.unobserve(entry.target);
  });
}, {
  threshold: 0.05,
  rootMargin: '0px 0px -15px 0px'
});

revealEls.forEach(el => observer.observe(el));


// ═══════════════════════════════════════
//  SMOOTH SCROLL per ancore interne
// ═══════════════════════════════════════
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
  anchor.addEventListener('click', e => {
    const href = anchor.getAttribute('href');
    const target = document.querySelector(href);
    
    if (target) {
      e.preventDefault();
      const offset = 65; // Altezza ottimizzata navbar
      const top = target.getBoundingClientRect().top + window.scrollY - offset;
      
      window.scrollTo({ top, behavior: 'smooth' });
    }
  });
});


// ═══════════════════════════════════════
//  ACTIVE NAVIGATION LINK TRACKING
// ═══════════════════════════════════════
const sections = document.querySelectorAll('section[id], div[id="transfer"], div[id="budget"]');
const swatches = document.querySelectorAll('.nav-links a');

const activeObserver = new IntersectionObserver(entries => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      const id = entry.target.id;
      
      swatches.forEach(swatch => {
        const targetHref = swatch.getAttribute('href');
        const isActive = targetHref === `#${id}` || 
                         (id === 'transfer' && targetHref === '#day1');
        
        if (isActive) {
          swatch.style.transform = 'scale(1.1) rotate(0deg) translateY(-4px)';
          swatch.style.boxShadow = '0 4px 12px rgba(0,0,0,0.18)';
        } else {
          swatch.style.transform = '';
          swatch.style.boxShadow = '';
        }
      });
    }
  });
}, { threshold: 0.25, rootMargin: '-60px 0px -40% 0px' });

sections.forEach(s => activeObserver.observe(s));


// ═══════════════════════════════════════
//  PARALLAX SOTTILE IMMAGINE HERO
// ═══════════════════════════════════════
const heroImg = document.querySelector('.hero-img');
if (heroImg) {
  let scrollTicking = false;
  window.addEventListener('scroll', () => {
    if (!scrollTicking) {
      requestAnimationFrame(() => {
        const y = window.scrollY;
        if (y < window.innerHeight * 1.3) {
          heroImg.style.transform = `scale(1.05) translateY(${y * 0.16}px)`;
        }
        scrollTicking = false;
      });
      scrollTicking = true;
    }
  }, { passive: true });
}

// ═══════════════════════════════════════
//  DISEGNO AUTOMATICO AL CARICAMENTO (Sprouting del titolo)
// ═══════════════════════════════════════
window.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.float-in').forEach((el, idx) => {
    setTimeout(() => {
      el.classList.add('is-visible');
    }, idx * 150);
  });
});
