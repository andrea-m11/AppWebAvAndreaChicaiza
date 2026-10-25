/* =====================================================
   PoliMusic · Reproductor global (HTML5 Audio)
   ===================================================== */
(function () {
    "use strict";

    var audio = document.getElementById("audio");
    if (!audio) return;

    var player = document.getElementById("player");
    var btnPlay = document.getElementById("btnPlay");
    var btnPrev = document.getElementById("btnPrev");
    var btnNext = document.getElementById("btnNext");
    var icoPlay = document.getElementById("icoPlay");
    var icoPause = document.getElementById("icoPause");
    var seek = document.getElementById("seek");
    var vol = document.getElementById("vol");
    var tCur = document.getElementById("tCur");
    var tDur = document.getElementById("tDur");
    var pImg = document.getElementById("pImg");
    var pTitle = document.getElementById("pTitle");
    var pArtist = document.getElementById("pArtist");

    var canciones = [];
    var indice = -1;
    var buscando = false;

    /* Ruta relativa a la raíz de la app (la emite layout.asp) */
    var BASE = (typeof RUTA_BASE !== "undefined") ? RUTA_BASE : "";

    /* ---------- Utilidades ---------- */
    function fmt(seg) {
        if (!isFinite(seg) || seg < 0) seg = 0;
        var m = Math.floor(seg / 60);
        var s = Math.floor(seg % 60);
        return m + ":" + (s < 10 ? "0" : "") + s;
    }

    function llenarRango(el, pct) {
        el.style.setProperty("--llenado", pct + "%");
    }

    /* ---------- Recolectar canciones de la página ---------- */
    function recolectar() {
        canciones = [];
        var filas = document.querySelectorAll(".song-row[data-src]");
        for (var i = 0; i < filas.length; i++) {
            canciones.push({
                id: filas[i].getAttribute("data-id"),
                nombre: filas[i].getAttribute("data-name"),
                artista: filas[i].getAttribute("data-artist") || "Artista desconocido",
                src: filas[i].getAttribute("data-src"),
                portada: filas[i].getAttribute("data-cover") || "",
                fila: filas[i]
            });
        }
    }

    function marcarFilaActiva() {
        var filas = document.querySelectorAll(".song-row.playing");
        for (var i = 0; i < filas.length; i++) filas[i].classList.remove("playing");
        if (indice >= 0 && canciones[indice]) {
            canciones[indice].fila.classList.add("playing");
        }
    }

    /* ---------- Contabilizar reproducción en el servidor ---------- */
    function contarReproduccion(id, celda) {
        fetch(BASE + "includes/reproducir.asp", {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: "id=" + encodeURIComponent(id)
        }).then(function (r) { return r.json(); }).then(function (d) {
            if (d && d.ok && celda) {
                var n = parseInt(celda.getAttribute("data-plays"), 10);
                if (isNaN(n)) n = 0;
                n += 1;
                celda.setAttribute("data-plays", n);
                celda.textContent = n;
            }
        }).catch(function () { /* silencioso: el contador es informativo */ });
    }

    function celdaPlays(fila) {
        return fila ? fila.querySelector(".plays-num") : null;
    }

    /* ---------- Reproducción ---------- */
    function cargar(i, reproducir) {
        if (i < 0 || i >= canciones.length) return;
        indice = i;
        var c = canciones[i];

        pImg.src = c.portada || (BASE + "SongFilesAndImages_resources/PlaylistImages/default.jpg");
        pTitle.textContent = c.nombre;
        pArtist.textContent = c.artista;

        if (audio.getAttribute("data-id") !== c.id) {
            audio.src = c.src;
            audio.setAttribute("data-id", c.id);
        }

        marcarFilaActiva();

        if (reproducir !== false) {
            var promesa = audio.play();
            if (promesa && typeof promesa.catch === "function") {
                promesa.catch(function () { /* el usuario no interactuó aún */ });
            }
        }
    }

    function alternar() {
        if (indice < 0 && canciones.length > 0) {
            cargar(0, true);
            return;
        }
        if (audio.paused) {
            var p = audio.play();
            if (p && typeof p.catch === "function") p.catch(function () {});
        } else {
            audio.pause();
        }
    }

    function siguiente() {
        if (canciones.length === 0) return;
        cargar((indice + 1) % canciones.length, true);
    }

    function anterior() {
        if (canciones.length === 0) return;
        var i = indice - 1;
        if (i < 0) i = canciones.length - 1;
        cargar(i, true);
    }

    /* ---------- Eventos del reproductor ---------- */
    btnPlay.addEventListener("click", alternar);
    btnNext.addEventListener("click", siguiente);
    btnPrev.addEventListener("click", anterior);

    audio.addEventListener("play", function () {
        icoPlay.style.display = "none";
        icoPause.style.display = "block";
        player.classList.add("girando");
        if (indice >= 0) marcarFilaActiva();
    });

    audio.addEventListener("pause", function () {
        icoPlay.style.display = "block";
        icoPause.style.display = "none";
        player.classList.remove("girando");
    });

    audio.addEventListener("ended", siguiente);

    audio.addEventListener("loadedmetadata", function () {
        tDur.textContent = fmt(audio.duration);
    });

    audio.addEventListener("timeupdate", function () {
        if (buscando) return;
        if (isFinite(audio.duration) && audio.duration > 0) {
            var pct = (audio.currentTime / audio.duration) * 100;
            seek.value = Math.round((audio.currentTime / audio.duration) * 1000);
            llenarRango(seek, pct);
        }
        tCur.textContent = fmt(audio.currentTime);
    });

    audio.addEventListener("error", function () {
        pArtist.textContent = "No se pudo cargar el audio de esta canción";
    });

    seek.addEventListener("input", function () {
        buscando = true;
        var pct = (seek.value / 1000) * 100;
        llenarRango(seek, pct);
        if (isFinite(audio.duration) && audio.duration > 0) {
            tCur.textContent = fmt((seek.value / 1000) * audio.duration);
        }
    });

    seek.addEventListener("change", function () {
        if (isFinite(audio.duration) && audio.duration > 0) {
            audio.currentTime = (seek.value / 1000) * audio.duration;
        }
        buscando = false;
    });

    vol.addEventListener("input", function () {
        audio.volume = vol.value / 100;
        llenarRango(vol, vol.value);
    });

    /* ---------- Botones de cada fila ---------- */
    function enlazarBotones() {
        recolectar();
        var botones = document.querySelectorAll(".song-row .row-play");
        for (var i = 0; i < botones.length; i++) {
            botones[i].addEventListener("click", function () {
                var fila = this.closest(".song-row");
                if (!fila) return;
                var id = fila.getAttribute("data-id");
                recolectar();
                var idx = -1;
                for (var j = 0; j < canciones.length; j++) {
                    if (canciones[j].id === id) { idx = j; break; }
                }
                if (idx < 0) return;

                var misma = (indice === idx);
                cargar(idx, !misma || audio.paused);
                if (misma && !audio.paused) return;

                contarReproduccion(canciones[idx].id, celdaPlays(canciones[idx].fila));
            });
        }
    }

    enlazarBotones();
    audio.volume = 0.8;
    llenarRango(vol, 80);
})();
