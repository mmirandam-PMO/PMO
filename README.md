# Tablero Portafolio de Proyectos 2027

Genera una diapositiva ejecutiva de PowerPoint (100 % editable) con el portafolio por clasificación,
responsable de cada proyecto y dos líneas de resumen en cada total.

## Uso
```bash
npm install
npm run build        # -> Portafolio_Proyectos_2027_Tablero.pptx
```

## Cómo actualizar
- **Datos** (proyectos, responsables, totales, resúmenes): edita `data.json`.
- **Diseño** (colores, tamaños, posiciones): edita `build.js`.
  - Paleta corporativa al inicio: Azul Profundo `#1B355D`, Naranja Legado `#E75301`, Rojo Cercanía `#983920`, Azul Horizonte `#7DACE3`.
  - `tint(hex, k)` mezcla un color con blanco para obtener los tonos claros.

## Estructura
```
data.json    # contenido del tablero
build.js     # diseño y generación del .pptx
package.json
```

> Las posiciones son fijas: si agregas más proyectos a una columna, ajusta las coordenadas `y` en `build.js`.
