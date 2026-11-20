/* SHERU BHAI — black screen; mouse reveals the money in the background,
   then everything fades back to black 2 seconds after the mouse stops. */

(() => {
  const money = document.querySelector(".money");
  const root = document.documentElement;
  const hint = document.getElementById("hint");

  if (!money || !root) return; // never crash on a stale/mismatched DOM

  const target = { x: window.innerWidth / 2, y: window.innerHeight / 2 };
  const cur = { x: target.x, y: target.y };

  let hideTimer = null;
  let moved = false;

  const HIDE_DELAY = 1200; // ms of stillness before the fade-out starts…
  // …plus the 0.8s CSS fade = money gone exactly 2s after last movement

  function onPointer(e) {
    target.x = e.clientX;
    target.y = e.clientY;

    if (!moved) {
      moved = true;
      hint.classList.add("is-hidden");
    }

    money.classList.add("is-on");
    clearTimeout(hideTimer);
    hideTimer = setTimeout(() => {
      money.classList.remove("is-on"); // fades back to black over 0.8s
    }, HIDE_DELAY);
  }

  // every flavor of movement — pointer, mouse and touch all count
  ["pointermove", "pointerdown", "mousemove", "touchstart", "touchmove"].forEach(
    (type) => window.addEventListener(type, onPointer, { passive: true })
  );

  // smooth spotlight follow
  function frame() {
    cur.x += (target.x - cur.x) * 0.14;
    cur.y += (target.y - cur.y) * 0.14;
    root.style.setProperty("--mx", cur.x.toFixed(1) + "px");
    root.style.setProperty("--my", cur.y.toFixed(1) + "px");
    requestAnimationFrame(frame);
  }

  requestAnimationFrame(frame);
})();
