/* SHERU BHAI — black screen; the cursor is a small brush that paints the
   money wallpaper onto the background. Everything disappears (fades back
   to full black) exactly 1 second after the cursor stops. */

(() => {
  const money = document.querySelector(".money");
  const hint = document.getElementById("hint");
  if (!money) return; // never crash on a stale/mismatched DOM

  const canvas = document.getElementById("veil");
  let ctx = null;
  try {
    ctx = canvas ? canvas.getContext("2d") : null;
  } catch (err) {
    ctx = null;
  }

  let w = 0, h = 0, dpr = 1;
  let last = null;        // last brush point
  let active = false;     // money currently showing
  let moved = false;      // has the user ever moved the cursor?
  let hideTimer = null;
  let clearTimer = null;
  let demoTimer = null;
  let demoRaf = 0;

  const HIDE_DELAY = 400; // stop -> start fading after 0.4s…
  const FADE_MS = 600;    // …css fade 0.6s = fully black 1s after last move

  function brushRadius() {
    return Math.max(36, Math.min(w, h) * 0.05); // a brush, not a big dot
  }

  function fillBlack() {
    if (!ctx) return;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.globalCompositeOperation = "source-over";
    ctx.fillStyle = "#000";
    ctx.fillRect(0, 0, w, h);
  }

  function resize() {
    if (!ctx || !canvas) return;
    const prev = document.createElement("canvas");
    if (canvas.width && canvas.height) {
      prev.width = canvas.width;
      prev.height = canvas.height;
      prev.getContext("2d").drawImage(canvas, 0, 0);
    }
    dpr = Math.min(window.devicePixelRatio || 1, 2);
    w = window.innerWidth;
    h = window.innerHeight;
    canvas.width = Math.round(w * dpr);
    canvas.height = Math.round(h * dpr);
    canvas.style.width = w + "px";
    canvas.style.height = h + "px";
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    if (prev.width) {
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.drawImage(prev, 0, 0, canvas.width, canvas.height);
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    } else {
      fillBlack();
    }
    last = null;
  }

  // one soft brush dab (destination-out punches a hole in the black veil)
  function dab(x, y) {
    if (!ctx) return;
    const r = brushRadius();
    const g = ctx.createRadialGradient(x, y, r * 0.5, x, y, r);
    g.addColorStop(0, "rgba(0,0,0,1)");
    g.addColorStop(0.75, "rgba(0,0,0,0.85)");
    g.addColorStop(1, "rgba(0,0,0,0)");
    ctx.globalCompositeOperation = "destination-out";
    ctx.fillStyle = g;
    ctx.beginPath();
    ctx.arc(x, y, r, 0, Math.PI * 2);
    ctx.fill();
    ctx.globalCompositeOperation = "source-over";
  }

  // continuous stroke: dab along the path so fast moves leave no gaps
  function stroke(x, y) {
    if (last) {
      const dx = x - last.x;
      const dy = y - last.y;
      const dist = Math.hypot(dx, dy);
      const steps = Math.max(1, Math.ceil(dist / (brushRadius() * 0.4)));
      for (let i = 1; i <= steps; i++) {
        dab(last.x + (dx * i) / steps, last.y + (dy * i) / steps);
      }
    } else {
      dab(x, y);
    }
    last = { x, y };
  }

  function show() {
    clearTimeout(hideTimer);
    clearTimeout(clearTimer);
    if (!active) {
      active = true;
      money.classList.add("is-on");
    }
  }

  function scheduleHide() {
    clearTimeout(hideTimer);
    hideTimer = setTimeout(() => {
      active = false;
      money.classList.remove("is-on"); // fades to black over 0.6s
      clearTimer = setTimeout(() => {
        if (!active) {
          fillBlack();  // once fully invisible, reset the canvas…
          last = null;  // …so the next stroke starts clean
        }
      }, FADE_MS + 60);
    }, HIDE_DELAY);
  }

  function paint(x, y) {
    stroke(x, y);
    show();
    scheduleHide();
  }

  function getPos(e) {
    if (e.touches && e.touches[0]) {
      return { x: e.touches[0].clientX, y: e.touches[0].clientY };
    }
    return { x: e.clientX, y: e.clientY };
  }

  function onMove(e) {
    const p = getPos(e);
    if (!moved) {
      moved = true;
      if (hint) hint.classList.add("is-hidden");
      clearTimeout(demoTimer);
      cancelAnimationFrame(demoRaf);
    }
    paint(p.x, p.y);
  }

  // every flavor of movement — pointer, mouse and touch all count
  ["pointermove", "pointerdown", "mousemove", "touchstart", "touchmove"].forEach(
    (type) => window.addEventListener(type, onMove, { passive: true })
  );

  window.addEventListener("resize", resize);

  // Safety net: if the environment never delivers a single pointer event,
  // run ONE gentle demo brush sweep. Any real input cancels it forever.
  demoTimer = setTimeout(() => {
    if (moved) return;
    const t0 = performance.now();
    const dur = 4200;
    (function sweep(now) {
      if (moved) return;
      const p = Math.min(1, (now - t0) / dur);
      paint(
        window.innerWidth * (0.25 + 0.5 * p),
        window.innerHeight * (0.45 + 0.14 * Math.sin(p * Math.PI * 2))
      );
      if (p < 1) demoRaf = requestAnimationFrame(sweep);
    })(t0);
  }, 6000);

  resize();
})();
