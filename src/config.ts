// Números de WhatsApp (código de país + número, solo dígitos). Al contactar se elige uno al azar.
export const WHATSAPP_NUMBERS = ['584146842498', '584246368716', '584246439327', '584146519332'];
// Valor por defecto para el HTML estático; el script lo reemplaza por uno aleatorio al hacer clic.
export const WHATSAPP_NUMBER = WHATSAPP_NUMBERS[0];

export const SITE_TITLE = 'Stock disponible | Ruedas Darío 4x4';
export const SITE_DESCRIPTION =
  'Catálogo de repuestos y accesorios 4x4 en stock. Elige lo que te interesa y consúltanos por WhatsApp.';

// Chips de vehículo: `terms` se busca (sin acentos, en minúsculas) dentro de la compatibilidad.
export const VEHICULOS = [
  { label: 'Hilux', terms: ['hilux'] },
  { label: '4Runner', terms: ['4runner'] },
  { label: 'Land Cruiser', terms: ['land cruiser', 'autana', 'roraima'] },
  { label: 'Fortuner', terms: ['fortuner'] },
  { label: 'Tacoma', terms: ['tacoma'] },
  { label: 'Tundra', terms: ['tundra'] },
  { label: 'Ranger', terms: ['ranger'] },
  { label: 'F150 / Super Duty', terms: ['f150', 'super duty'] },
  { label: 'Explorer', terms: ['explorer'] },
  { label: 'Grand Cherokee', terms: ['grand cherokee'] },
  { label: 'Silverado / Tahoe', terms: ['silverado', 'tahoe', 'avalanche', 'suburban'] },
  { label: 'L200', terms: ['l200'] },
  { label: 'Corolla', terms: ['corolla'] },
];
