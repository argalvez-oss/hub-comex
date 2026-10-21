// Configuración de Hub Comex.
// Completar antes de publicar (ver README.md, sección "Configuración en Azure").
window.HUB_CONFIG = {
  // Registro de la app en Microsoft Entra ID (Azure) → "Id. de aplicación (cliente)".
  // Registro propio "Hub Comex". PENDIENTE: un administrador debe aprobar sus permisos.
  // Alternativa sin administrador: usar el id de "Hub Recepción" (ya tiene aprobado Sites.ReadWrite.All),
  // agregarle las URI de redirección de Hub Comex y cambiar permisos a ['User.Read', 'Sites.ReadWrite.All'].
  clientId: '3abc78ee-335d-4231-a7d2-4dab471900a5',
  // Microsoft Entra ID → "Id. de directorio (inquilino)"
  tenantId: '8f9b210d-f5e5-404f-9fed-a0a827154105',

  // Permisos de Microsoft Graph que pide la app. Deben estar aprobados en el registro de clientId.
  permisos: ['User.Read', 'Files.ReadWrite'],

  // Únicos correos que pueden entrar, en minúsculas, y su rol.
  // "ventas": uso normal. "soporte": puede hacer todo, sus pruebas van a su propio OneDrive.
  correosPermitidos: {
    'ventas-sales@healthyfood.com.ec': 'ventas',
    'notificaciones@healthyfood.com.ec': 'soporte'
  },

  // Carpeta del OneDrive donde se crean las carpetas de cada despacho.
  carpetaRaiz: 'Despachos'
};
