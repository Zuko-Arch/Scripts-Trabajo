#target illustrator

function procesarArteLaserAI8() {
    if (app.documents.length === 0) {
        alert("No hay ningún documento abierto.");
        return;
    }

    var doc = app.activeDocument;

    if (doc.path == "") {
        alert("El archivo no ha sido guardado todavía.\nGuárdalo primero dentro de tu carpeta de trabajo para poder ubicar la carpeta anterior.");
        return;
    }

    if (doc.artboards.length !== 1) {
        alert("Proceso detenido: El documento tiene " + doc.artboards.length + " mesas de trabajo.\nDebe tener exactamente 1 mesa de trabajo.");
        return;
    }

    if (doc.layers.length !== 2) {
        alert("Proceso detenido: Para LÁSER deben existir exactamente 2 capas (arriba el arte, abajo la capa guía del cliente).\nActualmente hay " + doc.layers.length + " capa(s).");
        return;
    }

    var capaArte = doc.layers[0];
    var capaCliente = doc.layers[1];

    if (!capaArte.visible || capaArte.locked) {
        alert("Proceso detenido: La capa superior del arte ('" + capaArte.name + "') está oculta o bloqueada.");
        return;
    }

    if (capaArte.pageItems.length === 0) {
        alert("Proceso detenido: La capa superior del arte ('" + capaArte.name + "') está vacía.");
        return;
    }

    for (var a = 0; a < capaArte.pageItems.length; a++) {
        if (capaArte.pageItems[a].hidden || capaArte.pageItems[a].locked) {
            alert("Proceso detenido: Hay un elemento oculto o bloqueado dentro de la capa superior ('" + capaArte.name + "').");
            return;
        }
    }

    // 1. VALIDACIÓN DE TEXTOS SIN EXPANDIR EN LA CAPA DEL ARTE
    if (capaArte.textFrames.length > 0) {
        var textosSinExpandir = [];
        for (var t = 0; t < capaArte.textFrames.length; t++) {
            textosSinExpandir.push(capaArte.textFrames[t]);
        }
        doc.selection = textosSinExpandir;
        alert("Proceso detenido: Se detectaron " + textosSinExpandir.length + " texto(s) sin expandir en la capa '" + capaArte.name + "'.\nConviértelos a contornos (Texto > Crear contornos / Objeto > Expandir) antes de continuar.");
        return;
    }

    // 2. VALIDACIÓN DE SÍMBOLOS E IMÁGENES EN LA CAPA DEL ARTE (CUADRO ROJO PROPORCIONAL)
    var simbolosEImagenes = recolectarSimbolosEImagenesLaser(capaArte);
    if (simbolosEImagenes.length > 0) {
        encerrarObjetosEnCuadroRojoLaser(capaArte, simbolosEImagenes);
        alert("Proceso detenido: Se encontraron " + simbolosEImagenes.length + " símbolo(s) o imagen(es) (JPG/PNG) en '" + capaArte.name + "'.\nTodo debe ir únicamente en vectores y deben ser eliminados.\n\nAl cerrar este aviso, aparecerán encerrados en un cuadro rojo para que los ubiques y los borres.");
        return;
    }

    // 3. VALIDACIÓN DE OPACIDAD MENOR AL 100% O MODOS DE FUSIÓN
    var objetosTransparencia = recolectarTransparenciasLaser(capaArte);
    if (objetosTransparencia.length > 0) {
        doc.selection = objetosTransparencia;
        alert("Proceso detenido: Se detectaron " + objetosTransparencia.length + " elemento(s) con opacidad menor al 100% o con modo de fusión distinto a Normal en '" + capaArte.name + "'.\nPon la opacidad al 100% en modo Normal para evitar que se conviertan en imagen al exportar.");
        return;
    }

    // 4. VALIDACIÓN DE MÁSCARAS DE RECORTE Y DEGRADADOS
    var mascarasYDegradados = recolectarMascarasYDegradadosLaser(capaArte);
    if (mascarasYDegradados.length > 0) {
        doc.selection = mascarasYDegradados;
        alert("Proceso detenido: Se detectó una máscara de recorte o relleno degradado en '" + capaArte.name + "' (" + mascarasYDegradados.length + " elemento(s)) que debe ser eliminado o convertido a vector sólido.\n\nLos elementos implicados han sido seleccionados para que puedas eliminarlos.");
        return;
    }

    // 5. VALIDACIÓN DE OBJETOS INVISIBLES EN LA CAPA DEL ARTE
    var invisibles = detectarYMarcarInvisiblesLaser(capaArte);
    if (invisibles.length > 0) {
        doc.selection = invisibles;
        alert("Proceso detenido: Se encontraron " + invisibles.length + " objeto(s) invisible(s) (sin relleno ni trazo) en '" + capaArte.name + "' y no se puede continuar.\n\nAl cerrar este aviso, los trazos invisibles quedarán marcados con un borde de color rojo y seleccionados para que puedas verlos.");
        return;
    }

    // 6. VALIDACIÓN DE TRAZOS SIN EXPANDIR EN LA CAPA DEL ARTE
    var trazosLaser = [];
    for (var p = 0; p < capaArte.pathItems.length; p++) {
        if (capaArte.pathItems[p].stroked === true) {
            trazosLaser.push(capaArte.pathItems[p]);
        }
    }
    if (trazosLaser.length > 0) {
        doc.selection = trazosLaser;
        alert("Proceso detenido: Se detectaron " + trazosLaser.length + " elemento(s) con trazo sin expandir en la capa '" + capaArte.name + "'.");
        return;
    }

    // 7. VALIDACIÓN GEOMÉTRICA 1: VECTORES REALMENTE ENCIMADOS EN LA CAPA DEL ARTE
    var encimados = detectarVectoresEncimadosLaser(capaArte);
    if (encimados.length > 0) {
        doc.selection = encimados;
        var winEncimados = new Window("dialog", "Aviso: Vectores encimados detectados");
        winEncimados.orientation = "column";
        winEncimados.alignChildren = ["fill", "top"];
        winEncimados.margins = 20;
        winEncimados.spacing = 12;

        winEncimados.add("statictext", undefined, "Se detectaron vectores superpuestos uno encima de otro sin estar soldados en la capa '" + capaArte.name + "'.\nLos elementos implicados han sido seleccionados.", { multiline: true });

        var grpBtn1 = winEncimados.add("group");
        grpBtn1.alignment = ["right", "bottom"];
        grpBtn1.add("button", undefined, "Detener para arreglar", { name: "cancel" });
        grpBtn1.add("button", undefined, "Continuar sin importar este detalle", { name: "ok" });

        if (winEncimados.show() !== 1) {
            return;
        }
    }

    // 8. VALIDACIÓN GEOMÉTRICA 2: NODOS CRUZADOS / SUPERPUESTOS SOBRE EL RELLENO
    var puntosError = detectarNodosCruzadosLaser(capaArte);
    if (puntosError.length > 0) {
        var winDefectos = new Window("dialog", "Aviso: Nodos cruzados o encimados detectados");
        winDefectos.orientation = "column";
        winDefectos.alignChildren = ["fill", "top"];
        winDefectos.margins = 20;
        winDefectos.spacing = 12;

        winDefectos.add("statictext", undefined, "Se encontraron " + puntosError.length + " nodo(s) estirado(s) o cruzado(s) encima de otro nodo/relleno dentro del vector en '" + capaArte.name + "'.", { multiline: true });

        var grpBtn2 = winDefectos.add("group");
        grpBtn2.alignment = ["right", "bottom"];
        grpBtn2.add("button", undefined, "Arreglar (Marcar con cuadrado rojo)", { name: "cancel" });
        grpBtn2.add("button", undefined, "Continuar sin importar este error", { name: "ok" });

        if (winDefectos.show() !== 1) {
            marcarConCuadradosRojosLaser(doc, capaArte, puntosError);
            return;
        }
    }

    var estadoOriginalCapa2Visible = capaCliente.visible;
    capaCliente.locked = true;

    if (doc.documentColorSpace !== DocumentColorSpace.CMYK) {
        app.executeMenuCommand("doc-color-cmyk");
    }

    var negroPuroK100 = new CMYKColor();
    negroPuroK100.cyan = 0;
    negroPuroK100.magenta = 0;
    negroPuroK100.yellow = 0;
    negroPuroK100.black = 100;

    doc.selection = null;
    app.executeMenuCommand("selectall");
    aplicarNegroRecursivoLaser(doc.selection, negroPuroK100);

    doc.selection = null;
    app.executeMenuCommand("selectall");
    app.executeMenuCommand("group");

    limpiarMuestrasNoUsadasLaser(doc);

    capaCliente.visible = false;

    var nombreBase = doc.name.replace(/\.[^\.]+$/, '');
    var carpetaDestino = doc.path.parent;

    var archivoPNG = new File(carpetaDestino + "/" + nombreBase + ".png");
    var opcionesPNG = new ExportOptionsPNG24();
    opcionesPNG.antiAliasing = true;
    opcionesPNG.transparency = true;
    opcionesPNG.artBoardClipping = false;
    opcionesPNG.horizontalScale = 416.6667;
    opcionesPNG.verticalScale = 416.6667;

    doc.exportFile(archivoPNG, ExportType.PNG24, opcionesPNG);

    capaCliente.visible = estadoOriginalCapa2Visible;
    capaCliente.locked = true;

    var archivoAI = new File(carpetaDestino + "/" + nombreBase + ".ai");
    var opcionesAI = new IllustratorSaveOptions();
    opcionesAI.compatibility = Compatibility.ILLUSTRATOR8;
    opcionesAI.flattenOutput = OutputFlattening.PRESERVEAPPEARANCE;
    opcionesAI.compressed = true;

    doc.saveAs(archivoAI, opcionesAI);

    alert("¡Proceso LÁSER completado con éxito!\n\nArchivos guardados en:\n" + carpetaDestino.fsName + "\n- " + nombreBase + ".png (300 ppi - Solo arte)\n- " + nombreBase + ".ai (Illustrator 8 - Negro K100)");
}

// =========================================================================
// CÁLCULO DE ESCALA PROPORCIONAL AL TAMAÑO DEL ARTE (CAPA SUPERIOR)
// =========================================================================
function obtenerEscalaMarcaLaser(capaArte) {
    var minX = Infinity, maxY = -Infinity, maxX = -Infinity, minY = Infinity;
    for (var i = 0; i < capaArte.pageItems.length; i++) {
        var b = capaArte.pageItems[i].visibleBounds;
        if (b[0] < minX) minX = b[0];
        if (b[1] > maxY) maxY = b[1];
        if (b[2] > maxX) maxX = b[2];
        if (b[3] < minY) minY = b[3];
    }
    var anchoArte = (maxX > minX) ? (maxX - minX) : 20;
    var altoArte = (maxY > minY) ? (maxY - minY) : 20;
    var dimMax = Math.max(anchoArte, altoArte);
    if (dimMax <= 0) dimMax = 20;

    var tamanoCuadro = Math.max(1.5, Math.min(24, dimMax * 0.035));
    var grosorTrazo = Math.max(0.25, Math.min(2, dimMax * 0.0035));
    var margen = Math.max(0.5, Math.min(6, dimMax * 0.01));

    return {
        cuadro: tamanoCuadro,
        grosor: grosorTrazo,
        margen: margen
    };
}

// =========================================================================
// VALIDACIONES DE PREPRENSA
// =========================================================================
function recolectarSimbolosEImagenesLaser(contenedor) {
    var encontrados = [];
    function buscar(items) {
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (it.typename === "SymbolItem" || it.typename === "RasterItem" || it.typename === "PlacedItem") {
                encontrados.push(it);
            } else if (it.typename === "GroupItem") {
                buscar(it.pageItems);
            }
        }
    }
    buscar(contenedor.pageItems);
    return encontrados;
}

function encerrarObjetosEnCuadroRojoLaser(capa, objetos) {
    var escala = obtenerEscalaMarcaLaser(capa);
    var rojo = new RGBColor();
    rojo.red = 255;
    rojo.green = 0;
    rojo.blue = 0;

    var seleccionFinal = [];

    for (var i = 0; i < objetos.length; i++) {
        var obj = objetos[i];
        var b = obj.geometricBounds;
        var ancho = Math.max(Math.abs(b[2] - b[0]) + (escala.margen * 2), escala.cuadro);
        var alto = Math.max(Math.abs(b[1] - b[3]) + (escala.margen * 2), escala.cuadro);
        var centroX = (b[0] + b[2]) / 2;
        var centroY = (b[1] + b[3]) / 2;

        var rect = capa.pathItems.rectangle(
            centroY + (alto / 2),
            centroX - (ancho / 2),
            ancho,
            alto
        );
        rect.filled = false;
        rect.stroked = true;
        rect.strokeColor = rojo;
        rect.strokeWidth = escala.grosor;

        seleccionFinal.push(obj);
        seleccionFinal.push(rect);
    }
    app.activeDocument.selection = seleccionFinal;
}

function recolectarTransparenciasLaser(contenedor) {
    var encontrados = [];
    function buscar(items) {
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            var tieneOpacidadBaja = (it.opacity !== undefined && it.opacity < 99.9);
            var tieneModoFusion = (it.blendingMode !== undefined && it.blendingMode !== BlendModes.NORMAL);

            if (tieneOpacidadBaja || tieneModoFusion) {
                encontrados.push(it);
            } else if (it.typename === "GroupItem") {
                buscar(it.pageItems);
            } else if (it.typename === "CompoundPathItem") {
                buscar(it.pathItems);
            }
        }
    }
    buscar(contenedor.pageItems);
    return encontrados;
}

function recolectarMascarasYDegradadosLaser(contenedor) {
    var encontrados = [];
    function buscar(items) {
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (it.typename === "GroupItem") {
                if (it.clipped === true) {
                    encontrados.push(it);
                } else {
                    buscar(it.pageItems);
                }
            } else if (it.typename === "PathItem") {
                var esDegradado = (it.filled && it.fillColor && it.fillColor.typename === "GradientColor") ||
                                  (it.stroked && it.strokeColor && it.strokeColor.typename === "GradientColor");
                if (it.clipping === true || esDegradado) {
                    encontrados.push(it);
                }
            } else if (it.typename === "CompoundPathItem") {
                buscar(it.pathItems);
            }
        }
    }
    buscar(contenedor.pageItems);
    return encontrados;
}

function detectarYMarcarInvisiblesLaser(capa) {
    var todosPaths = [];
    recolectarTodosPathItemsLaser(capa, todosPaths);

    var escala = obtenerEscalaMarcaLaser(capa);
    var bordeRojo = new RGBColor();
    bordeRojo.red = 255;
    bordeRojo.green = 0;
    bordeRojo.blue = 0;

    var encontrados = [];
    for (var i = 0; i < todosPaths.length; i++) {
        var p = todosPaths[i];
        if (p.guides || p.clipping) continue;

        if (!p.filled && !p.stroked) {
            p.stroked = true;
            p.strokeColor = bordeRojo;
            p.strokeWidth = escala.grosor;
            encontrados.push(p);
        }
    }
    return encontrados;
}

// =========================================================================
// MOTOR DE DETECCIÓN GEOMÉTRICA: VECTORES ENCIMADOS Y NODOS CRUZADOS
// =========================================================================
function obtenerPiezasIndependientesLaser(contenedor, lista) {
    for (var i = 0; i < contenedor.pageItems.length; i++) {
        var it = contenedor.pageItems[i];
        if (it.typename === "PathItem" && it.filled) {
            lista.push(it);
        } else if (it.typename === "CompoundPathItem") {
            lista.push(it);
        } else if (it.typename === "GroupItem") {
            obtenerPiezasIndependientesLaser(it, lista);
        }
    }
}

function extraerPoligonosDePiezaLaser(pieza) {
    var poligonos = [];
    var paths = [];
    if (pieza.typename === "PathItem") {
        paths.push(pieza);
    } else if (pieza.typename === "CompoundPathItem") {
        for (var c = 0; c < pieza.pathItems.length; c++) {
            paths.push(pieza.pathItems[c]);
        }
    }

    for (var p = 0; p < paths.length; p++) {
        var path = paths[p];
        var pts = path.pathPoints;
        var n = pts ? pts.length : 0;
        if (n < 2) continue;
        var poly = [];
        var totalSegs = path.closed ? n : n - 1;
        for (var s = 0; s < totalSegs; s++) {
            var curr = pts[s];
            var next = pts[(s + 1) % n];
            var esCurva = (curr.rightDirection[0] !== curr.anchor[0] || curr.rightDirection[1] !== curr.anchor[1] ||
                           next.leftDirection[0] !== next.anchor[0] || next.leftDirection[1] !== next.anchor[1]);
            var pasos = esCurva ? 8 : 1;
            for (var step = 0; step < pasos; step++) {
                var t = step / pasos;
                var xy = esCurva
                    ? evaluarBezierLaser(curr.anchor, curr.rightDirection, next.leftDirection, next.anchor, t)
                    : curr.anchor;
                poly.push({ x: xy[0], y: xy[1] });
            }
        }
        if (!path.closed) {
            poly.push({ x: pts[n - 1].anchor[0], y: pts[n - 1].anchor[1] });
        }
        poligonos.push({ puntos: poly, cerrado: path.closed });
    }
    return poligonos;
}

function segmentosSeCruzanRealmenteLaser(a, b, c, d) {
    if (Math.max(a.x, b.x) < Math.min(c.x, d.x) - 0.1 || Math.min(a.x, b.x) > Math.max(c.x, d.x) + 0.1) return false;
    if (Math.max(a.y, b.y) < Math.min(c.y, d.y) - 0.1 || Math.min(a.y, b.y) > Math.max(c.y, d.y) + 0.1) return false;

    var d1x = b.x - a.x, d1y = b.y - a.y;
    var d2x = d.x - c.x, d2y = d.y - c.y;
    var denom = d1x * d2y - d1y * d2x;
    if (Math.abs(denom) < 0.0001) return false;

    var t = ((c.x - a.x) * d2y - (c.y - a.y) * d2x) / denom;
    var u = ((c.x - a.x) * d1y - (c.y - a.y) * d1x) / denom;

    return (t > 0.02 && t < 0.98 && u > 0.02 && u < 0.98);
}

function obtenerPuntoDeCruceLaser(a, b, c, d) {
    if (Math.max(a.x, b.x) < Math.min(c.x, d.x) || Math.min(a.x, b.x) > Math.max(c.x, d.x)) return null;
    if (Math.max(a.y, b.y) < Math.min(c.y, d.y) || Math.min(a.y, b.y) > Math.max(c.y, d.y)) return null;

    var d1x = b.x - a.x, d1y = b.y - a.y;
    var d2x = d.x - c.x, d2y = d.y - c.y;
    var denom = d1x * d2y - d1y * d2x;
    if (Math.abs(denom) < 0.0001) return null;

    var t = ((c.x - a.x) * d2y - (c.y - a.y) * d2x) / denom;
    var u = ((c.x - a.x) * d1y - (c.y - a.y) * d1x) / denom;

    if (t > 0.005 && t < 0.995 && u > 0.005 && u < 0.995) {
        return { x: a.x + t * d1x, y: a.y + t * d1y };
    }
    return null;
}

function puntoDentroDePoligonoLaser(pt, poly) {
    var dentro = false;
    var n = poly.length;
    for (var i = 0, j = n - 1; i < n; j = i++) {
        var xi = poly[i].x, yi = poly[i].y;
        var xj = poly[j].x, yj = poly[j].y;
        var intersecta = ((yi > pt.y) !== (yj > pt.y)) &&
                         (pt.x < (xj - xi) * (pt.y - yi) / ((yj - yi) || 0.00001) + xi);
        if (intersecta) dentro = !dentro;
    }
    return dentro;
}

function puntoDentroDePiezaLaser(pt, poligonosPieza) {
    var conteo = 0;
    for (var i = 0; i < poligonosPieza.length; i++) {
        if (poligonosPieza[i].cerrado && puntoDentroDePoligonoLaser(pt, poligonosPieza[i].puntos)) {
            conteo++;
        }
    }
    return (conteo % 2) === 1;
}

function piezasSeSuperponenRealmenteLaser(polisA, polisB) {
    for (var i = 0; i < polisA.length; i++) {
        var pA = polisA[i].puntos;
        var lenA = polisA[i].cerrado ? pA.length : pA.length - 1;
        for (var sa = 0; sa < lenA; sa++) {
            var a1 = pA[sa];
            var a2 = pA[(sa + 1) % pA.length];
            for (var j = 0; j < polisB.length; j++) {
                var pB = polisB[j].puntos;
                var lenB = polisB[j].cerrado ? pB.length : pB.length - 1;
                for (var sb = 0; sb < lenB; sb++) {
                    var b1 = pB[sb];
                    var b2 = pB[(sb + 1) % pB.length];
                    if (segmentosSeCruzanRealmenteLaser(a1, a2, b1, b2)) {
                        return true;
                    }
                }
            }
        }
    }

    if (polisA.length > 0 && polisA[0].puntos.length > 0) {
        if (puntoDentroDePiezaLaser(polisA[0].puntos[0], polisB)) return true;
    }

    if (polisB.length > 0 && polisB[0].puntos.length > 0) {
        if (puntoDentroDePiezaLaser(polisB[0].puntos[0], polisA)) return true;
    }

    return false;
}

function detectarVectoresEncimadosLaser(capa) {
    var piezas = [];
    obtenerPiezasIndependientesLaser(capa, piezas);
    if (piezas.length < 2) return [];

    var poligonosPorPieza = [];
    for (var p = 0; p < piezas.length; p++) {
        poligonosPorPieza.push(extraerPoligonosDePiezaLaser(piezas[p]));
    }

    var conflictivos = [];
    var yaAgregado = {};

    for (var i = 0; i < piezas.length; i++) {
        var b1 = piezas[i].geometricBounds;
        for (var j = i + 1; j < piezas.length; j++) {
            var b2 = piezas[j].geometricBounds;
            var solapaCaja = (b1[0] < b2[2] - 0.2) && (b1[2] > b2[0] + 0.2) &&
                             (b1[1] > b2[3] + 0.2) && (b1[3] < b2[1] - 0.2);
            if (!solapaCaja) continue;

            if (piezasSeSuperponenRealmenteLaser(poligonosPorPieza[i], poligonosPorPieza[j])) {
                if (!yaAgregado[i]) { conflictivos.push(piezas[i]); yaAgregado[i] = true; }
                if (!yaAgregado[j]) { conflictivos.push(piezas[j]); yaAgregado[j] = true; }
            }
        }
    }
    return conflictivos;
}

function recolectarTodosPathItemsLaser(contenedor, lista) {
    for (var i = 0; i < contenedor.pageItems.length; i++) {
        var it = contenedor.pageItems[i];
        if (it.typename === "PathItem") {
            lista.push(it);
        } else if (it.typename === "CompoundPathItem") {
            for (var c = 0; c < it.pathItems.length; c++) {
                lista.push(it.pathItems[c]);
            }
        } else if (it.typename === "GroupItem") {
            recolectarTodosPathItemsLaser(it, lista);
        }
    }
}

function agregarMarcaUnicaLaser(puntos, x, y, distMinSq) {
    for (var r = 0; r < puntos.length; r++) {
        var dx = puntos[r].x - x;
        var dy = puntos[r].y - y;
        if ((dx * dx + dy * dy) < distMinSq) return;
    }
    puntos.push({ x: x, y: y });
}

function evaluarBezierLaser(p0, p1, p2, p3, t) {
    var u = 1 - t;
    var tt = t * t;
    var uu = u * u;
    var uuu = uu * u;
    var ttt = tt * t;
    return [
        uuu * p0[0] + 3 * uu * t * p1[0] + 3 * u * tt * p2[0] + ttt * p3[0],
        uuu * p0[1] + 3 * uu * t * p1[1] + 3 * u * tt * p2[1] + ttt * p3[1]
    ];
}

function detectarNodosCruzadosLaser(capa) {
    var puntos = [];
    var todosPaths = [];
    recolectarTodosPathItemsLaser(capa, todosPaths);

    var escala = obtenerEscalaMarcaLaser(capa);
    var distMinSq = (escala.cuadro * 0.6) * (escala.cuadro * 0.6);

    for (var i = 0; i < todosPaths.length; i++) {
        var p = todosPaths[i];
        var pts = p.pathPoints;
        var n = pts ? pts.length : 0;
        if (n < 3) continue;

        for (var v = 0; v < n; v++) {
            var vx = pts[v].anchor[0];
            var vy = pts[v].anchor[1];

            var nextIdx = (v + 1) % n;
            if (p.closed || v < n - 1) {
                var nx = pts[nextIdx].anchor[0];
                var ny = pts[nextIdx].anchor[1];
                if (Math.abs(vx - nx) < 0.01 && Math.abs(vy - ny) < 0.01) {
                    agregarMarcaUnicaLaser(puntos, vx, vy, distMinSq);
                    continue;
                }
            }

            if (p.closed || (v > 0 && v < n - 1)) {
                var prevPt = pts[(v - 1 + n) % n].anchor;
                var nextPt = pts[(v + 1) % n].anchor;
                var v1x = prevPt[0] - vx, v1y = prevPt[1] - vy;
                var v2x = nextPt[0] - vx, v2y = nextPt[1] - vy;
                var mag1 = Math.sqrt(v1x * v1x + v1y * v1y);
                var mag2 = Math.sqrt(v2x * v2x + v2y * v2y);
                if (mag1 > 0.05 && mag2 > 0.05) {
                    var cosAng = (v1x * v2x + v1y * v2y) / (mag1 * mag2);
                    if (cosAng > 0.996) {
                        agregarMarcaUnicaLaser(puntos, vx, vy, distMinSq);
                    }
                }
            }
        }

        var poly = [];
        var totalSegs = p.closed ? n : n - 1;
        for (var s = 0; s < totalSegs; s++) {
            var curr = pts[s];
            var next = pts[(s + 1) % n];
            var esCurva = (curr.rightDirection[0] !== curr.anchor[0] || curr.rightDirection[1] !== curr.anchor[1] ||
                           next.leftDirection[0] !== next.anchor[0] || next.leftDirection[1] !== next.anchor[1]);
            var pasos = esCurva ? 8 : 1;
            for (var step = 0; step < pasos; step++) {
                var t = step / pasos;
                var xy = esCurva
                    ? evaluarBezierLaser(curr.anchor, curr.rightDirection, next.leftDirection, next.anchor, t)
                    : curr.anchor;
                poly.push({ x: xy[0], y: xy[1] });
            }
        }
        if (!p.closed) {
            poly.push({ x: pts[n - 1].anchor[0], y: pts[n - 1].anchor[1] });
        }

        var mLen = poly.length;
        var numSegsPoly = p.closed ? mLen : mLen - 1;

        for (var segA = 0; segA < numSegsPoly; segA++) {
            var a1 = poly[segA];
            var a2 = poly[(segA + 1) % mLen];
            for (var segB = segA + 2; segB < numSegsPoly; segB++) {
                if (p.closed && segA === 0 && segB === numSegsPoly - 1) continue;
                var b1 = poly[segB];
                var b2 = poly[(segB + 1) % mLen];

                var cruce = obtenerPuntoDeCruceLaser(a1, a2, b1, b2);
                if (cruce !== null) {
                    agregarMarcaUnicaLaser(puntos, cruce.x, cruce.y, distMinSq);
                }
            }
        }
    }
    return puntos;
}

function marcarConCuadradosRojosLaser(doc, capa, puntos) {
    var escala = obtenerEscalaMarcaLaser(capa);
    var rojo = new RGBColor();
    rojo.red = 255;
    rojo.green = 0;
    rojo.blue = 0;

    var marcas = [];
    var tamanoCuadro = escala.cuadro;

    for (var i = 0; i < puntos.length; i++) {
        var pt = puntos[i];
        var rect = capa.pathItems.rectangle(
            pt.y + (tamanoCuadro / 2),
            pt.x - (tamanoCuadro / 2),
            tamanoCuadro,
            tamanoCuadro
        );
        rect.filled = false;
        rect.stroked = true;
        rect.strokeColor = rojo;
        rect.strokeWidth = escala.grosor;
        marcas.push(rect);
    }

    doc.selection = marcas;
    alert("Se encerró con un cuadrado rojo la ubicación del nodo cruzado/encimado.\nCorrige el nodo, borra el cuadrado rojo y vuelve a ejecutar el script.");
}

function aplicarNegroRecursivoLaser(items, colorNegro) {
    for (var i = 0; i < items.length; i++) {
        var obj = items[i];
        if (obj.typename === "PathItem") {
            if (obj.filled === true) obj.fillColor = colorNegro;
        } else if (obj.typename === "CompoundPathItem") {
            for (var c = 0; c < obj.pathItems.length; c++) {
                if (obj.pathItems[c].filled === true) obj.pathItems[c].fillColor = colorNegro;
            }
        } else if (obj.typename === "GroupItem") {
            aplicarNegroRecursivoLaser(obj.pageItems, colorNegro);
        }
    }
}

function limpiarMuestrasNoUsadasLaser(doc) {
    for (var g = doc.swatchGroups.length - 1; g >= 1; g--) {
        try { doc.swatchGroups[g].remove(); } catch (e) {}
    }

    var nombresEnUso = {};
    recolectarColoresEnUsoLaser(doc.pageItems, nombresEnUso);

    for (var s = doc.swatches.length - 1; s >= 0; s--) {
        var sw = doc.swatches[s];
        var nombre = sw.name;
        if (nombre === "[Ninguno]" || nombre === "[None]" || nombre === "[Registro]" || nombre === "[Registration]") {
            continue;
        }
        if (!nombresEnUso[nombre]) {
            try { sw.remove(); } catch (e) {}
        }
    }
}

function recolectarColoresEnUsoLaser(items, mapa) {
    for (var i = 0; i < items.length; i++) {
        var obj = items[i];
        if (obj.typename === "PathItem") {
            if (obj.filled && obj.fillColor) registrarColorLaser(obj.fillColor, mapa);
            if (obj.stroked && obj.strokeColor) registrarColorLaser(obj.strokeColor, mapa);
        } else if (obj.typename === "CompoundPathItem") {
            recolectarColoresEnUsoLaser(obj.pathItems, mapa);
        } else if (obj.typename === "GroupItem") {
            recolectarColoresEnUsoLaser(obj.pageItems, mapa);
        }
    }
}

function registrarColorLaser(color, mapa) {
    if (color.typename === "SpotColor" && color.spot) {
        mapa[color.spot.name] = true;
    } else if (color.typename === "PatternColor" && color.pattern) {
        mapa[color.pattern.name] = true;
    } else if (color.typename === "GradientColor" && color.gradient) {
        mapa[color.gradient.name] = true;
    }
}

procesarArteLaserAI8();