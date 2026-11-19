// Se inyecta en una copia temporal de un HTML antes de abrirlo con Edge sin
// ventana. Sirve para páginas que generan su contenido con JavaScript, como
// las "clases interactivas" de bloques que se pasan con «Siguiente».
//
// Qué hace, una vez cargada la página:
//   1. Si hay un botón de avanzar (Siguiente, Next, →...), vuelve al principio
//      y recorre todos los pasos, guardando la parte de la página que cambia.
//   2. En cada paso, pulsa la primera opción de cada pregunta para que la
//      página muestre cuál es la correcta y su explicación.
//   3. Deja todo lo recogido en <main id="__convmd__"> para que Python lo lea.
(function () {
  "use strict";
  var MAX_PASOS = 400;
  var AVANZAR = /siguiente|next|continuar|adelante|→|›|»|▶/i;
  var RETROCEDER = /anterior|previous|prev\b|atrás|volver|←|‹|«|◀/i;

  function esperar(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function texto(el) { return (el.innerText || el.textContent || "").trim(); }
  function visible(el) {
    var s = getComputedStyle(el);
    return s.display !== "none" && s.visibility !== "hidden" && el.getClientRects().length > 0;
  }
  function usable(b) { return b && !b.disabled && b.getAttribute("aria-disabled") !== "true" && visible(b); }

  function botones() {
    return Array.prototype.slice.call(
      document.querySelectorAll("button, [role=button], a.btn, a.button, input[type=button]"));
  }
  function buscar(patron, contrario) {
    var lista = botones().filter(function (b) {
      var t = texto(b) || b.value || b.getAttribute("aria-label") || b.title || "";
      return t.length <= 30 && patron.test(t) && !contrario.test(t);
    });
    return lista.filter(usable)[0] || lista[0] || null;
  }

  function foto() {
    var m = new Map();
    document.body.querySelectorAll("*").forEach(function (el) { m.set(el, el.innerHTML); });
    return m;
  }
  // Baja desde <body> hacia el hijo que cambió y que tiene casi todo el texto.
  function contenedorQueCambia(antes) {
    var actual = document.body;
    for (;;) {
      var largo = texto(actual).length, siguiente = null;
      for (var i = 0; i < actual.children.length; i++) {
        var h = actual.children[i];
        var cambio = !antes.has(h) || antes.get(h) !== h.innerHTML;
        if (cambio && texto(h).length >= 0.6 * largo) { siguiente = h; break; }
      }
      if (!siguiente) return actual;
      actual = siguiente;
    }
  }

  // Grupos de opciones: un elemento con 2 o más botones hijos que no navegan.
  function gruposDeOpciones(raiz) {
    var grupos = [];
    raiz.querySelectorAll("*").forEach(function (el) {
      var hijos = Array.prototype.filter.call(el.children, function (h) {
        return h.matches("button, [role=button], input[type=button]") &&
          !AVANZAR.test(texto(h)) && !RETROCEDER.test(texto(h));
      });
      if (hijos.length >= 2) grupos.push(hijos);
    });
    return grupos;
  }
  async function revelarRespuestas(raiz) {
    var grupos = gruposDeOpciones(raiz);
    for (var i = 0; i < grupos.length; i++) {
      var primero = grupos[i][0];
      if (usable(primero)) {
        try { primero.click(); } catch (e) {}
        await esperar(30);
      }
    }
  }

  async function recorrer() {
    var salida = [];
    var siguiente = buscar(AVANZAR, RETROCEDER);
    if (!siguiente || !usable(siguiente)) {
      await revelarRespuestas(document.body);
      return null;                      // página normal: vale el DOM tal cual
    }
    // Volver al principio por si la página recordaba otra posición.
    for (var n = 0; n < MAX_PASOS; n++) {
      var atras = buscar(RETROCEDER, AVANZAR);
      if (!usable(atras)) break;
      var antesAtras = document.body.innerHTML;
      atras.click();
      await esperar(40);
      if (document.body.innerHTML === antesAtras) break;
    }
    // Averiguar qué parte de la página cambia al avanzar.
    var antes = foto();
    var copia = document.body.cloneNode(true);
    siguiente = buscar(AVANZAR, RETROCEDER);
    siguiente.click();
    await esperar(60);
    var caja = contenedorQueCambia(antes);
    if (caja === document.body) return null;
    var atras2 = buscar(RETROCEDER, AVANZAR);
    if (usable(atras2)) {
      atras2.click();
      await esperar(60);
    } else {
      // Sin «Anterior»: el primer paso se toma de la copia previa.
      var ruta = [], el = caja;
      while (el !== document.body) { ruta.unshift(Array.prototype.indexOf.call(el.parentNode.children, el)); el = el.parentNode; }
      var enCopia = copia;
      ruta.forEach(function (i) { enCopia = enCopia && enCopia.children[i]; });
      if (enCopia) salida.push(enCopia.innerHTML);
    }
    var vistos = new Set(salida);
    for (var paso = 0; paso < MAX_PASOS; paso++) {
      caja = document.contains(caja) ? caja : contenedorQueCambia(new Map());
      await revelarRespuestas(caja);
      var html = caja.innerHTML;
      if (!vistos.has(html)) { salida.push(html); vistos.add(html); }
      siguiente = buscar(AVANZAR, RETROCEDER);
      if (!usable(siguiente)) break;
      var previo = caja.innerHTML;
      siguiente.click();
      await esperar(60);
      if (caja.innerHTML === previo) break;
    }
    return salida;
  }

  window.addEventListener("load", function () {
    setTimeout(async function () {
      var pasos = null;
      try { pasos = await recorrer(); } catch (e) { pasos = null; }
      var main = document.createElement("main");
      main.id = "__convmd__";
      if (pasos && pasos.length) {
        main.setAttribute("data-pasos", String(pasos.length));
        main.innerHTML = pasos.map(function (h) { return "<section>" + h + "</section>"; }).join("<hr>");
      } else {
        main.innerHTML = document.body.innerHTML;
      }
      document.body.innerHTML = "";
      document.body.appendChild(main);
    }, 300);
  });
})();
