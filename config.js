// Configuración de Hub Comex.
// Completar antes de publicar (ver README.md, sección "Configuración en Azure").
window.HUB_CONFIG = {
  // 'firebase' : ingreso con correo y contraseña; registros y fotos en la nube (Firebase), iguales en todos los dispositivos.
  // 'prueba'   : se elige Ventas o Soporte al ingresar y todo se guarda solo en el dispositivo.
  // 'microsoft': ingreso con la cuenta de Microsoft y fotos en OneDrive (requiere los permisos aprobados).
  modo: 'firebase',

  // false: sin cuentas, se elige Ventas o Soporte al ingresar; la app se identifica de forma anónima ante Firebase
  //        (Authentication → Anónimo habilitado; reglas: request.auth != null).
  // true : correo y contraseña creados en Firebase → Authentication (y reglas solo para esos correos).
  firebaseCuentas: false,

  // Proyecto de Firebase "hub-comex" (consola → Configuración del proyecto → Tus apps → Hub Comex).
  // No es una contraseña: es la dirección pública del proyecto. La seguridad la dan las reglas de Firestore.
  firebase: {
    apiKey: 'AIzaSyCYDX-w_gD3XJ7BAdi5U0m8UFnP3RYSnwc',
    authDomain: 'hub-comex.firebaseapp.com',
    projectId: 'hub-comex',
    storageBucket: 'hub-comex.firebasestorage.app',
    messagingSenderId: '130916289619',
    appId: '1:130916289619:web:90518a5a48496eda6acaa6'
  },

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

  // Carpeta del OneDrive donde va todo: adentro quedan "Despachos" y "Traslados", con una carpeta por registro.
  carpetaRaiz: 'Hub Comex',

  // Listas desplegables: se cargan desde catalogos.js (generado del Excel "LISTA DESPLEGABLE APP COMEX.xlsx"
  // con herramientas/excel-a-catalogos.py). Esto de abajo solo se usa si catalogos.js no existe.
  catalogos: {
    clientes: [],
    destinos: [],
    lugaresCarga: [],
    transportistas: [],
    navieras: [],
    productos: [
      // { cliente: 'Goya Foods Inc.', producto: 'Pulpa de mango 14 oz', codigo: 'PM-014' },
    ]
  },

  // Datos del encabezado del informe en Word (los asigna Calidad / el sistema de gestión).
  informe: {
    codigo: 'HF-LOG-___',
    revision: '___',
    empresa: 'Healthy Food Marcalman Ecuador S.A.S.'
  }
};
