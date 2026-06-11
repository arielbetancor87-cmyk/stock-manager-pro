# Cómo generar el APK de Venta Directa

## Lo que necesitás instalar primero

1. **Node.js** (si no lo tenés)
   - Descargá de https://nodejs.org → versión LTS
   - Instalá normalmente

2. **Android Studio**
   - Descargá de https://developer.android.com/studio
   - Durante la instalación, asegurate de instalar:
     - Android SDK
     - Android SDK Platform-Tools

---

## Pasos para generar el APK

### 1. Cloná o descargá el proyecto
```
git clone https://github.com/arielbetancor87-cmyk/stock-manager-pro
cd stock-manager-pro
```

### 2. Reemplazá el package.json
Copiá el archivo `package_capacitor.json` y renombralo a `package.json`
(reemplazando el existente)

### 3. Copiá el capacitor.config.json
Pegá el archivo `capacitor.config.json` en la raíz del proyecto

### 4. Instalá dependencias
```
npm install
```

### 5. Inicializá Capacitor Android
```
npx cap add android
```

### 6. Construí la app
```
npm run build
npx cap sync android
```

### 7. Abrí en Android Studio
```
npx cap open android
```

### 8. Generá el APK en Android Studio
- En Android Studio: **Build → Build Bundle(s) / APK(s) → Build APK(s)**
- El APK se genera en: `android/app/build/outputs/apk/debug/app-debug.apk`

### 9. Instalá en el celular
- Conectá el celular por USB con **depuración USB activada**
- O copiá el `.apk` al celular y abrilo para instalar

---

## Resultado
- Ícono rojo "Venta Directa" en la pantalla de inicio
- Abre como app nativa sin browser
- Splash screen rojo al abrir
- Barra de estado roja

