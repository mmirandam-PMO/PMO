# Portafolio de Proyectos 2027

Genera la presentación completa en PowerPoint (100 % editable), 7 diapositivas:
1. Tablero del portafolio
2. Presupuesto por clasificación (sin total general)
3-7. Timeline + presupuesto, una diapositiva por clasificación (una fila por proyecto)

## Uso
    npm install
    npm run build        # -> Portafolio_Proyectos_2027.pptx

## Cómo actualizar
- Datos: edita data.json
  - Tablero (diapositiva 1): claves minorista, mayorista, regulatorios, hub, carryover.
  - Timeline (diapositivas 3-7): clave "clasificaciones". Cada fila es:
      [proyecto, línea, responsable, plan_ini, ejec_ini, ejec_fin, peso?]
    Los meses van de 0 (Ene) a 11 (Dic). "peso" es opcional: 0 para programas TBD que no cuentan como proyecto.
- Diseño (colores, tamaños, posiciones): edita build.js. La paleta está al inicio.

## Archivos
    build.js
    data.json
    package.json
    .gitignore   (contenido: node_modules/ y *.pptx, una por línea)
