/* SHERU BHAI — cursor reveals the wealth through the black veil */

(() => {
  const canvas = document.getElementById("veil");
  const ctx = canvas.getContext("2d");
  const root = document.documentElement;
  const hint = document.getElementById("hint");

  let dpr = Math.min(window.devicePixelRatio || 1, 2);
  let w = 0, h = 0;
  let last = null;          // last erase point (css px)
  let moved = false;
  let lastMove = performance.now();

  // target = where the reveal should be; cur = smoothed position
  const target = { x: window.innerWidth / 2, y: window.innerHeight * 0.45 };
  const cur = { x: target.x, y: target.y };

  function fillVeil() {
    ctx.globalCompositeOperation = "source-over";
    ctx.fillStyle = "#000";
    ctx.fillRect(0, 0, w, h);
  }

  function resize() {
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
      fillVeil();
    }
    last = null;
  }

  // erase a soft circular hole (destination-out punches through the black)
  function erase(x, y, radius) {
    const g = ctx.createRadialGradient(x, y, radius * 0.45, x, y, radius);
    g.addColorStop(0, "rgba(0,0,0,1)");
    g.addColorStop(0.7, "rgba(0,0,0,0.85)");
    g.addColorStop(1, "rgba(0,0,0,0)");
    ctx.globalCompositeOperation = "destination-out";
    ctx.fillStyle = g;
    ctx.beginPath();
    ctx.arc(x, y, radius, 0, Math.PI * 2);
    ctx.fill();
    ctx.globalCompositeOperation = "source-over";
  }

  function radius() {
    return Math.max(150, Math.min(w, h) * 0.24); // big juicy reveal
  }

  function onPointer(e) {
    target.x = e.clientX;
    target.y = e.clientY;
    lastMove = performance.now();
    if (!moved) {
      moved = true;
      hint.classList.add("is-hidden");
    }
  }

  window.addEventListener("pointermove", onPointer, { passive: true });
  window.addEventListener("pointerdown", onPointer, { passive: true });
  window.addEventListener("resize", resize);

  function frame(now) {
    // idle: let the reveal drift on its own (figure-8 path) so touch /
    // no-cursor visitors still see the money
    if (now - lastMove > 2600) {
      target.x = w * (0.5 + 0.34 * Math.sin(now / 3200));
      target.y = h * (0.48 + 0.30 * Math.sin(now / 2100));
      if (!moved && now > 3200) hint.classList.add("is-hidden");
    }

    // smooth follow
    cur.x += (target.x - cur.x) * 0.08;
    cur.y += (target.y - cur.y) * 0.08;

    const r = radius();
    if (last) {
      // interpolate so fast sweeps don't leave gaps
      const dx = cur.x - last.x;
      const dy = cur.y - last.y;
      const dist = Math.hypot(dx, dy);
      const steps = Math.max(1, Math.ceil(dist / (r * 0.35)));
      for (let i = 1; i <= steps; i++) {
        erase(last.x + (dx * i) / steps, last.y + (dy * i) / steps, r);
      }
    } else {
      erase(cur.x, cur.y, r);
    }
    last = { x: cur.x, y: cur.y };

    // feed CSS vars (parallax + gradient focus)
    root.style.setProperty("--mx", cur.x.toFixed(1) + "px");
    root.style.setProperty("--my", cur.y.toFixed(1) + "px");

    requestAnimationFrame(frame);
  }

  resize();
  requestAnimationFrame(frame);
})();
