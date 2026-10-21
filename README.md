# Hub Comex

App del área de Ventas para fotografiar despachos y guardar las fotos directo en OneDrive.

- **Ingreso:** cuenta Microsoft de la empresa. Solo entran los correos listados en `config.js`. La sesión queda guardada en el dispositivo.
- **Fotos:** se guardan en el OneDrive de quien ingresa, en `Despachos/AAAA-MM-DD - <nombre del despacho>/Foto_001.jpg`, `Foto_002.jpg`…
- **Sin señal:** las fotos aceptadas quedan guardadas en el celular y se suben solas cuando vuelve la conexión, siempre que la app esté abierta.
- **Registro:** el archivo `_registro-hub-comex.json`, dentro de `OneDrive/Despachos`. **No borrarlo ni moverlo.**
- **Inicio de sesión (temporal):** usa el registro de Azure de *Hub Recepción*, que ya tiene aprobado `Sites.ReadWrite.All`. Cuando un administrador apruebe el registro propio *Hub Comex*, cambiar `clientId` y `permisos` en `config.js` (ahí están las instrucciones).

Archivos: `index.html`, `styles.css`, `app.js` (toda la lógica), `config.js` (datos a completar), `sw.js` (permite instalar la app y abrirla sin señal), `manifest.webmanifest`, `icons/`.

---

## 1. Configuración en Azure (una sola vez)

1. Entrar a <https://portal.azure.com> con la cuenta de la empresa → **Microsoft Entra ID** → **Registros de aplicaciones** → **Nuevo registro**.
2. Completar:
   - **Nombre:** `Hub Comex`
   - **Tipos de cuenta admitidos:** *Solo las cuentas de este directorio organizativo*
   - **URI de redirección:** tipo **Aplicación de página única (SPA)** → `https://sisgeshealthy.github.io/hub-comex/`
3. Tocar **Registrar**. En la pantalla *Información general* copiar:
   - **Id. de aplicación (cliente)** → `clientId` en `config.js`
   - **Id. de directorio (inquilino)** → `tenantId` en `config.js`
4. **Autenticación** → en *Aplicación de página única* agregar también `http://localhost:5500/` (para probar en la computadora) → **Guardar**.
5. **Permisos de API** → **Agregar un permiso** → **Microsoft Graph** → **Permisos delegados** → marcar `Files.ReadWrite` (además de `User.Read`, que ya viene) → **Agregar permisos**.

La primera vez que ventas ingrese, Microsoft le pedirá aceptar esos permisos. Si en cambio aparece *"Se necesita la aprobación del administrador"*, la empresa no deja que los usuarios los acepten. En ese caso, un administrador de Microsoft 365 debe tocar **Conceder consentimiento de administrador** en esa misma pantalla.

## 2. Completar `config.js`

- `clientId` y `tenantId`: los del paso anterior.
- `correosPermitidos`: el correo de Ventas con rol `'ventas'` y el de soporte con rol `'soporte'`, en minúsculas.

## 3. Publicar en GitHub Pages

1. En la organización **sisgeshealthy** crear el repositorio `hub-comex`.
2. Subir todos los archivos de esta carpeta a la rama `main`.
3. En el repositorio: **Settings → Pages → Source: Deploy from a branch → main / (root)** → **Save**.
4. En uno o dos minutos la app queda en `https://sisgeshealthy.github.io/hub-comex/`.

**Al publicar cambios:** subir el número de `VERSION` en `sw.js` (por ejemplo `hc-v1` → `hc-v2`). Así los celulares que ya tienen la app instalada descargan la versión nueva.

## 4. Probar en la computadora

Desde esta carpeta:

```bash
npx serve -l 5500
```

Abrir `http://localhost:5500/`. Tiene que ser `localhost` y no `127.0.0.1`, porque así está registrado en Azure.

## 5. Instalar en el celular

Abrir el enlace en **Chrome** → menú **⋮** → **Instalar aplicación** (o *Agregar a la pantalla principal*). Son los mismos pasos que en el manual de HUB Recepciones.
