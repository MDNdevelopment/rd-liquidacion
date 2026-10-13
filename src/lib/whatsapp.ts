import { WHATSAPP_NUMBER } from '../config';

export interface Producto {
  id: string;
  nombre: string;
  marca?: string;
  detalle?: string;
  compatibilidad: string[];
}

export function productoLabel(p: Producto): string {
  return p.marca ? `${p.nombre} (${p.marca})` : p.nombre;
}

export function waLink(text: string): string {
  return `https://wa.me/${WHATSAPP_NUMBER}?text=${encodeURIComponent(text)}`;
}

export interface ItemCarrito { label: string; qty: number }

export function carritoMensaje(items: ItemCarrito[], vehiculo = ''): string {
  const lineas = items.map((i) => `• ${i.qty} x *${i.label}*`).join('\n');
  const veh = vehiculo ? `\nEs para mi *${vehiculo}*.` : '';
  return `Hola Ruedas Darío 👋 Vi el catálogo y me interesan estos productos:\n${lineas}${veh}\n¿Me pueden dar precio y disponibilidad?`;
}

export const MENSAJE_GENERAL =
  'Hola Ruedas Darío 👋 Vi el catálogo de stock y quiero hacer una consulta.';
