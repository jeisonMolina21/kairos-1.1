// Definir un array para almacenar las 10 edades
let edades = [];
let cantidadMenores = 0;
let cantidadMayores = 0;
let cantidadAdultosMayores = 0;
let edadMasBaja = 121; // Inicializar con un valor superior al máximo permitido
let edadMasAlta = 0;
let sumaEdades = 0;

const TAMANO_GRUPO = 10;
const EDAD_MINIMA_VALIDA = 1;
const EDAD_MAXIMA_VALIDA = 120;
const ADULTO_MAYOR_UMBRAL = 60;

// Solicitar y validar las edades
for (let i = 0; i < TAMANO_GRUPO; i++) {
  let edadValida = false;

  while (!edadValida) {
    let entrada = prompt(`Ingrese la edad de la persona ${i + 1} (entre 1 y 120):`);
    let edad = parseInt(entrada);

    if (isNaN(edad) || edad < EDAD_MINIMA_VALIDA || edad > EDAD_MAXIMA_VALIDA) {
      alert("Error: La edad debe ser un número entre 1 y 120.");
    } else {
      edades.push(edad);
      edadValida = true;
    }
  }
}

// Procesar las edades para obtener los resultados
for (let i = 0; i < edades.length; i++) {
  const edad = edades[i];
  sumaEdades += edad;

  // Actualizar edad más baja y alta
  if (edad < edadMasBaja) edadMasBaja = edad;
  if (edad > edadMasAlta) edadMasAlta = edad;

  // Clasificar edades (SIN doble conteo)
  if (edad < 18) {
    cantidadMenores++;
  } else if (edad >= 18 && edad < ADULTO_MAYOR_UMBRAL) {
    cantidadMayores++;
  } else {
    cantidadAdultosMayores++;
  }
}

// Calcular el promedio
let promedioEdades = sumaEdades / TAMANO_GRUPO;

// Mostrar resultados
console.log(`Resultados del grupo de ${TAMANO_GRUPO} personas:`);
console.log(`Menores de edad: ${cantidadMenores}`);
console.log(`Mayores de edad (18-59): ${cantidadMayores}`);
console.log(`Adultos mayores (>= 60 años): ${cantidadAdultosMayores}`);
console.log(`Edad más baja: ${edadMasBaja}`);
console.log(`Edad más alta: ${edadMasAlta}`);
console.log(`Promedio de edades: ${promedioEdades.toFixed(2)}`);
