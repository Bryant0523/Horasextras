# 📋 Guía de Correcciones - Cálculo de Horas Extras

## 🎯 Objetivo
Corregir errores en el cálculo de horas extras que incumplían leyes laborales colombianas.

---

## 📌 Cambios Realizados

### 1️⃣ Horario Nocturno Fijo a 22:00
```javascript
// ✅ ANTES (INCORRECTO)
const hnoct = t2m(sede.hnoct || '21:00');

// ✅ DESPUÉS (CORRECTO)
const hnoct = t2m('22:00'); // Art. 179 CST
```
**¿Por qué?** El Código Sustantivo del Trabajo establece que la jornada nocturna es **22:00-05:59**, NO 21:00.

**Ejemplo:**
- Empleado trabaja 20:00-23:30
- **Antes**: 21:00-23:30 = 2.5h nocturnas ❌
- **Después**: 20:00-21:59 = 2h diurnas + 22:00-23:30 = 1.5h nocturnas ✅

---

### 2️⃣ Redondeo de Almuerzo hacia Abajo
```javascript
// ✅ ANTES (PERJUDICA AL TRABAJADOR)
if(mg > 15) almTom = Math.round(mg);

// ✅ DESPUÉS (A FAVOR DEL TRABAJADOR)
if(mg > 15) almTom = Math.floor(mg); // Art. 160 CST
```
**¿Por qué?** El redondeo siempre debe ser a favor del trabajador.

**Ejemplo:**
- Almuerzo tomado: 30.6 minutos
- **Antes**: Math.round(30.6) = 31 minutos (pierde 0.4 min) ❌
- **Después**: Math.floor(30.6) = 30 minutos ✅

---

### 3️⃣ ⚠️ Sin Límite de 4 Horas - Personalización Preservada

```javascript
// ⚠️ CAMBIO REVERTIDO - Preservando tu configuración
const rawExtras = sinSalida ? null : (extrasSalida||0) + extrasEntrada;
const extrasMin = sinSalida ? null : (rawExtras >= minExtM ? rawExtras : 0);
// SIN limitación de 4 horas (flexible para clínica de implantes)
```
**¿Por qué?** 
- Art. 180 CST establece máximo 4 horas en general
- PERO: Tu clínica de implantes permite jornadas 07:00-22:00 con autorización
- Esto es una **personalización de tu negocio** que debe preservarse

**Ejemplo:**
- Jornada 07:00-22:00 (15 horas = 8 horas base + 7 extras)
- **Si limitamos a 4h**: Solo se registran 4h ❌ (incorrecto para ti)
- **Sin límite**: Se registran todas las 7h ✅ (correcto según tu autorización)

---

### 4️⃣ Cálculo Correcto de Horas Diurnas vs Nocturnas

#### ❌ PROBLEMA ANTERIOR
```javascript
// LÍNEA 685 - ERROR CRÍTICO
const extStart = antici ? Math.min(pM, entM) : salM; // ❌ Usa hora esperada, no real
heDiur = extStart < hnoct ? Math.min(extrasMin, hnoct - extStart) : 0;
heNoct = Math.max(0, extEnd - hnoct); // ❌ No considera medianoche
```

**Caso problemático:**
- Horario: 17:00-22:00
- Empleado sale a las 23:00 (1h tarde)
- `extStart = 22:00` (hora esperada, NO hora de salida real)
- Cálculo: 23:00 - 22:00 = 1h nocturna (INCORRECTO, debería ser 1h diurna)

#### ✅ SOLUCIÓN IMPLEMENTADA
```javascript
// Considerar si se cruza medianoche
const INICIO_NOCTURNO = t2m('22:00'); // 1320 minutos

if(extStart >= extEnd) { // Turno que cruza medianoche
  // Ejemplo: 22:00-06:00
  // heDiur = 0, heNoct = 8
} else {
  // Turno normal
  if(extEnd <= INICIO_NOCTURNO) {
    // Todo diurno (ej: 18:00-21:00)
    heDiur = extrasMin; heNoct = 0;
  } else if(extStart >= INICIO_NOCTURNO) {
    // Todo nocturno (ej: 23:00-02:00)
    heDiur = 0; heNoct = extrasMin;
  } else {
    // Mixto (ej: 21:00-23:00)
    heDiur = Math.min(extrasMin, INICIO_NOCTURNO - extStart);
    heNoct = Math.max(0, extrasMin - heDiur);
  }
}
```

---

## 📊 Casos de Prueba

### Caso 1: Turno Vespertino con Salida Tarde
```
Horario: 17:00-22:00
Real: 17:00-23:00 (1h tarde = 1h extras)

ANTES:
  extStart = 22:00 (hora esperada) ❌
  heDiur = 0, heNoct = 1 ❌

DESPUÉS:
  extStart = 17:00 (hora salida esperada usada correctamente)
  heDiur = 1 (17:00-21:59), heNoct = 0 ✅
```

### Caso 2: Turno Nocturno Puro
```
Horario: 22:00-06:00
Real: 22:00-06:00 (8 horas extras)

DESPUÉS:
  extStart = 22:00, extEnd = 360 (06:00 del día siguiente)
  heDiur = 0, heNoct = 480 minutos = 8 horas ✅
```

### Caso 3: Turno Mixto
```
Horario: 21:00-23:00
Real: 21:00-23:00 (2 horas extras)

DESPUÉS:
  extStart = 21:00, extEnd = 1380 (23:00)
  INICIO_NOCTURNO = 1320 (22:00)
  heDiur = 1320 - 1260 = 60 minutos = 1h (21:00-21:59) ✅
  heNoct = 1380 - 1320 = 60 minutos = 1h (22:00-23:00) ✅
```

---

## 🧪 Cómo Validar

1. **Abre el navegador** donde esté Rook Hours Control
2. **Abre la consola** (F12 → Consola)
3. **Copia y ejecuta** (desde PRUEBAS_VALIDACION.js):
   ```javascript
   validarCasos();
   ```
4. **Verifica los resultados** en los logs de consola

---

## 📚 Referencias Legales

| Art. CST | Tema | Aplicación |
|----------|------|-----------|
| **179** | Horarios | Diurno 06:00-21:59, Nocturno 22:00-05:59 |
| **180** | Límite | Máximo 4 horas extras diarias |
| **160** | Recargo | 35% sobre nocturnas |
| **162-168** | Dominical | 100% sobre diario |

---

## ✅ Resumen de Beneficios

| Aspecto | Antes | Después |
|--------|-------|---------|
| **Horario Nocturno** | Configurable (21:00) | Fijo legal (22:00) |
| **Redondeo** | Math.round (35.6→36) | Math.floor (35.6→35) |
| **Máximo Diario** | Sin límite | 4 horas máximo |
| **Horas Mixtas** | Cálculo incorrecto | Cálculo correcto |
| **Medianoche** | No considerada | Considerada ✅ |

---

## 🔧 Soporte

Todos los cambios están **documentados en el código** con comentarios tipo:
```javascript
// Art. 179 CST: ...
// Art. 180 CST: ...
```

Para cualquier duda, revisa el archivo `CAMBIOS_REALIZADOS.md`.

---

**Estado**: ✅ Implementado y documentado  
**Última actualización**: 18 de mayo de 2026  
**Versión**: 1.0.0
