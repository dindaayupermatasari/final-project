import { useEffect } from "react";

// Elemen [data-s] bergeser vertikal relatif terhadap induknya saat scroll.
export default function useParallax() {
  useEffect(() => {
    if (matchMedia("(prefers-reduced-motion: reduce)").matches) return;
    const els = [...document.querySelectorAll("[data-s]")];
    let ticking = false;
    const run = () => {
      ticking = false;
      const vh = innerHeight;
      els.forEach((el) => {
        const r = el.parentElement.getBoundingClientRect();
        if (r.bottom < -300 || r.top > vh + 300) return;
        const d = (r.top + r.height / 2 - vh / 2) * parseFloat(el.dataset.s);
        el.style.transform = `translate3d(0,${d.toFixed(1)}px,0) rotate(${el.dataset.rot || 0}deg)`;
      });
    };
    const onScroll = () => { if (!ticking) { ticking = true; requestAnimationFrame(run); } };
    addEventListener("scroll", onScroll, { passive: true });
    addEventListener("resize", run);
    run();
    return () => { removeEventListener("scroll", onScroll); removeEventListener("resize", run); };
  }, []);
}
