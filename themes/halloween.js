/* Au Temps du Pain — thème Halloween : éléments décoratifs animés (lune, fantôme, feuilles, bulles...).
   Chargé par applyTheme() (index.html) uniquement quand THEME = halloween.
   Ne fait qu'ajouter / déplacer des éléments décoratifs (tous marqués data-ha) ;
   unmount() remet le DOM exactement comme avant. Les prix, le panier et les commandes ne sont pas touchés. */
(function(){
  const DECO = "themes/halloween/decorations/";
  const HOLD_AT = 11.3;   /* s : dernière image avec le message complet, juste avant le fondu au noir */
  const HOLD_MS = 30000;  /* durée de la pause sur le message */
  let mounted = false;
  let undo = [];

  function rnd(seed){ let s = seed; return () => { s = (s * 9301 + 49297) % 233280; return s / 233280; }; }
  /* Balises <i>/<div>/<img> uniquement : la règle .placeholder span du site ne doit pas les atteindre. */
  function el(tag, css, attrs){
    const n = document.createElement(tag);
    n.setAttribute("data-ha", "");
    if (css) n.style.cssText = css;
    if (attrs) for (const k in attrs) n.setAttribute(k, attrs[k]);
    return n;
  }
  function deco(file, css){ return el("img", "position:absolute;pointer-events:none;" + css, { src: DECO + file, alt: "", "aria-hidden": "true" }); }
  function add(parent, child, before){ before ? parent.insertBefore(child, before) : parent.appendChild(child); undo.push(() => child.remove()); return child; }
  function hidden(n){ n.setAttribute("aria-hidden", "true"); return n; }

  function ghost(live){
    const g = hidden(el("div", "position:absolute;top:-46px;right:14%;width:40px;height:50px;pointer-events:none;" + (live ? "animation:pa-ghost 5s ease-in-out infinite" : "")));
    const body = el("div", "position:absolute;inset:0;background:rgba(244,237,224,.88);border-radius:50% 50% 18% 18%/60% 60% 22% 22%;box-shadow:0 0 24px rgba(244,237,224,.35)");
    [12, 23].forEach(l => body.appendChild(el("i", "position:absolute;top:15px;left:" + l + "px;width:5px;height:7px;border-radius:50%;background:#2B1E16")));
    body.appendChild(el("i", "position:absolute;top:28px;left:17px;width:6px;height:6px;border-radius:50%;background:#2B1E16"));
    g.appendChild(body);
    return g;
  }

  function falling(){
    const r = rnd(21);
    const box = hidden(el("div", "position:absolute;inset:-40px -32px 0;overflow:hidden;pointer-events:none;z-index:3"));
    for (let i = 0; i < 10; i++) {
      const w = 14 + r() * 16;
      box.appendChild(el("img", "position:absolute;top:-30px;left:" + (r() * 100) + "%;width:" + w + "px;opacity:.9;--dx:" + ((r() - .5) * 160) + "px;--rot:" + ((r() > .5 ? 1 : -1) * (240 + r() * 360)) + "deg;animation:pa-fall " + (11 + r() * 9).toFixed(1) + "s linear " + (-r() * 18).toFixed(1) + "s infinite", { src: DECO + "leaf.svg", alt: "" }));
    }
    return box;
  }

  function fournil(live){
    const ink = "#3A2A1F";
    const f = hidden(el("div", "position:relative;width:130px;height:96px;flex:none"));
    for (let i = 0; i < 3; i++) f.appendChild(el("i", "position:absolute;left:86px;bottom:76px;width:14px;height:14px;border-radius:50%;background:rgba(216,203,179,.45);filter:blur(2px);" + (live ? "animation:pa-smoke 3.6s ease-out " + (i * 1.2) + "s infinite;opacity:0" : "opacity:.3")));
    f.appendChild(el("i", "position:absolute;left:84px;bottom:52px;width:14px;height:26px;background:" + ink));
    f.appendChild(el("i", "position:absolute;left:0;bottom:42px;width:130px;height:34px;background:" + ink + ";clip-path:polygon(50% 0,100% 100%,0 100%)"));
    f.appendChild(el("i", "position:absolute;left:12px;bottom:0;width:106px;height:44px;background:" + ink));
    [26, 92].forEach(l => f.appendChild(el("i", "position:absolute;bottom:16px;left:" + l + "px;width:12px;height:14px;background:#E37A2C;box-shadow:0 0 12px 2px rgba(227,122,44,.7);" + (live ? "animation:pa-flicker 2s ease-in-out " + (l / 40) + "s infinite" : ""))));
    f.appendChild(el("i", "position:absolute;left:56px;bottom:0;width:18px;height:26px;border-radius:9px 9px 0 0;background:#17100B;box-shadow:inset 0 0 0 1.5px rgba(227,122,44,.5)"));
    return f;
  }

  function steam(){
    const s = hidden(el("div", "position:absolute;left:0;right:0;bottom:34%;height:0;pointer-events:none;z-index:2"));
    [38, 50, 62].forEach((x, i) => s.appendChild(el("i", "position:absolute;left:" + x + "%;bottom:0;width:12px;height:52px;border-radius:50%;background:linear-gradient(to top,transparent,rgba(244,237,224,.45),transparent);filter:blur(3px);animation:pa-steam 2.4s ease-out " + (i * 0.7) + "s infinite")));
    return s;
  }

  /* Vidéo « Halloween Bannière » (800 x 800, 13 s, muette) à la place de la photo du hero.
     Elle s'arrête sur le message complet à HOLD_AT pendant HOLD_MS, puis joue le fondu final et recommence.
     Si le visiteur préfère moins de mouvement, on n'affiche que l'image fixe. */
  function scene(live){
    const box = el("div", "position:absolute;inset:0;overflow:hidden;z-index:1;background:#0E0B14");
    box.className = "ha-scene";
    box.setAttribute("data-ha-scene", "");
    if (live) {
      const v = el("video", "position:absolute;inset:0;width:100%;height:100%;object-fit:cover;display:block", {
        src: "themes/halloween/video/banniere.mp4", poster: "themes/halloween/video/banniere-poster.jpg",
        autoplay: "", muted: "", playsinline: "", preload: "auto", "aria-label": "Pâton de levain, nuit d’Halloween"
      });
      v.muted = true; /* l'attribut seul ne suffit pas toujours pour l'autoplay */
      box.appendChild(v);
      const play = () => { const p = v.play(); if (p && p.catch) p.catch(() => {}); };
      let timer = 0, resumed = false;
      const check = () => {
        if (resumed || v.paused || v.currentTime < HOLD_AT) return;
        resumed = true; v.pause();
        timer = setTimeout(play, HOLD_MS);
      };
      if (v.requestVideoFrameCallback) { const tick = () => { check(); if (v.isConnected) v.requestVideoFrameCallback(tick); }; v.requestVideoFrameCallback(tick); }
      else v.addEventListener("timeupdate", check);
      v.addEventListener("ended", () => { resumed = false; v.currentTime = 0; play(); });
      undo.push(() => clearTimeout(timer));
      play();
    } else {
      box.appendChild(el("img", "position:absolute;inset:0;width:100%;height:100%;object-fit:cover;display:block", { src: "themes/halloween/video/banniere-poster.jpg", alt: "Pâton de levain, nuit d’Halloween" }));
    }
    return box;
  }

  function mount(){
    if (mounted) return;
    mounted = true;
    undo = [];
    let live = true;
    try { live = !window.matchMedia("(prefers-reduced-motion: reduce)").matches; } catch (e) {}

    const hero = document.querySelector(".hero");
    const ph = document.getElementById("heroPhoto");
    if (hero && ph) {
      /* La scène « Halloween Pâte » remplace la photo du hero (même format 5/4 que la maquette, 1350 x 1080). */
      add(ph, scene(live), ph.firstChild);
      if (live) add(hero, falling());
      const txt = hero.firstElementChild;
      if (txt) add(txt, ghost(live), txt.firstChild);
    }

    const h2 = document.querySelector("#boutique .section-head h2");
    if (h2) add(h2, el("img", "height:34px;margin-left:14px;vertical-align:middle;transform-origin:20% 95%;" + (live ? "animation:pa-sway 4s ease-in-out infinite" : ""), { src: DECO + "wheat.svg", alt: "", "aria-hidden": "true" }));

    const fb = document.querySelector("footer .foot-bottom");
    if (fb) {
      const wrap = el("span");
      while (fb.firstChild) wrap.appendChild(fb.firstChild);
      fb.appendChild(wrap);
      undo.push(() => { while (wrap.firstChild) fb.insertBefore(wrap.firstChild, wrap); wrap.remove(); });
      add(fb, fournil(live));
    }

    /* Vapeur au survol d'une carte produit (les cartes sont recréées à chaque filtre : on délègue). */
    const over = e => {
      const card = e.target.closest && e.target.closest(".product");
      const img = card && card.querySelector(".product__img");
      if (live && img && !img.querySelector("[data-ha-steam]")) { const s = steam(); s.setAttribute("data-ha-steam", ""); img.appendChild(s); }
    };
    const out = e => {
      const card = e.target.closest && e.target.closest(".product");
      if (card && !card.contains(e.relatedTarget)) card.querySelectorAll("[data-ha-steam]").forEach(n => n.remove());
    };
    document.addEventListener("mouseover", over);
    document.addEventListener("mouseout", out);
    undo.push(() => { document.removeEventListener("mouseover", over); document.removeEventListener("mouseout", out); document.querySelectorAll("[data-ha-steam]").forEach(n => n.remove()); });
  }

  function unmount(){
    if (!mounted) return;
    undo.reverse().forEach(f => f());
    undo = [];
    mounted = false;
  }

  (window.SITE_THEME_MODULES = window.SITE_THEME_MODULES || {}).halloween = { mount, unmount };
})();
