import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function generateMessage(row: {
  ninos: string[];
  rit?: string;
  mes: string;
  remitente: string;
}) {
  const { ninos, rit, mes, remitente } = row;

  // Rule 1: Introduction
  let message = `Estimada/o, buenos días. Soy ${remitente.trim()} del programa DCE San Bernardo 2. Le escribo a petición del Tribunal de Familia. \n\n`;

  // Rule 2: Motivo & Causa
  const isPlural = ninos.length > 1;
  const childrenList = isPlural 
    ? ninos.map(n => `- ${n.trim()}`).join('\n')
    : ninos[0]?.trim() || '';

  const namesDisplay = isPlural ? 'los niños' : 'el niño';

  if (rit && rit.trim() !== '') {
    message += `Necesitamos tomar contacto con el adulto responsable para realizar la evaluación de ${namesDisplay} de la causa ${rit.trim()}.\n\n`;
    if (isPlural) {
      message += `Los niños son:\n${childrenList}\n\n`;
    }
  } else {
    if (isPlural) {
      message += `Necesitamos tomar contacto con el adulto responsable para realizar la evaluación de los niños:\n${childrenList}\n\n`;
    } else {
      message += `Necesitamos tomar contacto con el adulto responsable para realizar la evaluación del niño ${childrenList}. `;
    }
  }

  // Rule 3: Formato de validación EXACTO al pedido
  message += `Para poder continuar, necesitamos confirmar con quién estamos hablando. \n\nPor favor indicar:\n1-\tnombre completo \n2-\tqué relación tiene con ${namesDisplay} (padre, madre, abuelo/a, tío/a u otro).\n3-\tdirección actual. \n\n`;

  // Rule 4: Cierre
  message += `Quedamos atentos a su respuesta para poder iniciar este proceso, el cual comenzará durante el mes de ${mes.trim()}. Muchas gracias.`;

  return message;
}
