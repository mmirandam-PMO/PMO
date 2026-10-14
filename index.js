import fs from 'fs';

try {
  // Leer el archivo JSON con los datos mapeados de la imagen
  const rawData = fs.readFileSync('./proyectos.json', 'utf-8');
  const data = JSON.parse(rawData);

  console.log(`=== ${data.portafolio} ===`);
  console.log(`Programas mapeados: ${data.metricas_globales.programas}`);
  console.log(`Total de Iniciativas estimadas: ${data.metricas_globales.proyectos_totales}\n`);

  // Recorrer los programas y listar sus proyectos asociados
  data.programas.forEach(programa => {
    console.log(`🔹 Programa: ${programa.nombre}`);
    
    // Si contiene categorías directas (ej: Nuevos proyectos, Carryover)
    if (programa.categorias) {
      programa.categorias.forEach(cat => {
        cat.proyectos.forEach(p => {
          console.log(`   [ ] ${p.nombre} -> Líder: ${p.responsable || 'Sin asignar'}`);
        });
      });
    }
    
    // Si contiene subprogramas complejos (ej: Banca Mayorista)
    if (programa.subprogramas) {
      programa.subprogramas.forEach(sub => {
        console.log(`   🔸 Subprograma: ${sub.nombre}`);
        sub.proyectos.forEach(p => {
          const lider = p.responsable || (sub.responsables ? sub.responsables.join(' / ') : 'TBD');
          console.log(`       [ ] ${p.nombre} -> Líder: ${lider}`);
        });
      });
    }
  });

} catch (error) {
  console.error("Error procesando los datos del portafolio:", error.message);
}
