import gsap from 'gsap';
import { ScrollTrigger } from 'gsap/ScrollTrigger';

/**
 * Enterprise GSAP & ScrollTrigger Motion Kernel
 * Strictly client-side, leak-free animation orchestration.
 */
export function initGsapMotion() {
  if (typeof window === 'undefined') return;

  gsap.registerPlugin(ScrollTrigger);

  // Stagger reveal for elements with .gsap-reveal
  const revealElements = document.querySelectorAll('.gsap-reveal');
  if (revealElements.length > 0) {
    gsap.from(revealElements, {
      scrollTrigger: {
        trigger: revealElements[0],
        start: 'top 85%',
        toggleActions: 'play none none reverse'
      },
      y: 24,
      opacity: 0,
      stagger: 0.08,
      duration: 0.6,
      ease: 'power2.out',
      clearProps: 'all'
    });
  }
}

export { gsap, ScrollTrigger };
