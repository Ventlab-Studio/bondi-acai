# Evaluación Financiera — Bondi Açaí

## Metodología y Propuesta al Inversionista

**Fecha:** Abril 2026
**Preparado para:** Negociación con inversionista(s)
**Archivo complementario:** `Bondi_Acai_Evaluacion_Financiera.xlsx`

---

## 1. Resumen Ejecutivo de la Evaluación

El proyecto Bondi Açaí es **financieramente viable** bajo los tres escenarios evaluados. Todos los indicadores de evaluación financiera arrojan resultados positivos:

| Indicador | Conservador | Moderado | Optimista |
|-----------|-------------|----------|-----------|
| VPN | $9.8M COP | $334M COP | $812M COP |
| TIR anual | 36.5% | 550.5% | 1,739% |
| RBC | 1.18 | 8.67 | 21.30 |
| Payback | >18 meses | 10.9 meses | 5.5 meses |

**Conclusión:** Incluso en el escenario más pesimista, el proyecto crea valor (VPN > 0) y supera el costo de oportunidad del capital (TIR 36.5% > WACC 21.8%).

---

## 2. Metodología Utilizada

### 2.1 Modelo CAPM para el Costo del Capital

Al ser 100% financiado con equity (sin deuda), el WACC es igual al Costo del Equity (Ke):

```
Ke = Rf + β × (Rm - Rf) + Prima por tamaño

Ke = 10.5% + 0.9 × 7.0% + 5.0% = 21.8% anual
```

| Componente | Valor | Justificación |
|------------|-------|---------------|
| Rf (tasa libre de riesgo) | 10.5% | TES Colombia 10 años (2026) |
| β (beta del sector) | 0.90 | Sector alimentos / food service, sin apalancamiento |
| Rm - Rf (prima de mercado) | 7.0% | Prima de riesgo del mercado accionario colombiano |
| Prima por tamaño/startup | 5.0% | Riesgo adicional por ser micro-empresa nueva |

**Nota:** Si el proyecto tuviera deuda, el WACC sería menor por el escudo fiscal. Al usar 100% equity, estamos siendo conservadores en la tasa de descuento.

### 2.2 Flujo de Caja Libre (FCF)

Se calculó el FCF mensual como:

```
Ingresos
- Costos variables (COGS 32% + Packaging 3% + Delivery 1.5%)
= Utilidad Bruta (margen ~63.5%)
- Costos fijos operativos
= EBITDA
- Impuesto (Régimen Simple: 3.5% de ingresos brutos)
= Flujo de Caja Libre
```

La depreciación se incluyó en el P&L pero se suma de vuelta al FCF por ser gasto no efectivo.

### 2.3 Régimen Tributario

Se asume **Régimen Simple de Tributación** (RST), óptimo para este tipo de negocio:
- Tarifa consolidada: 3.5% sobre ingresos brutos
- Incluye: Impuesto de renta + ICA + Impuesto al consumo
- Más simple administrativamente que el Régimen Ordinario
- Aplica para SAS con ingresos < 100,000 UVT/año (~$5,000M COP)

### 2.4 Valor Terminal

Para capturar el valor del negocio más allá del horizonte de 18 meses:

```
Valor Terminal = EBITDA mensual (mes 18) × 12 × 3.5x

Donde 3.5x es el múltiplo EV/EBITDA, conservador para food service local
(rango industria: 2-5x)
```

### 2.5 Tres Escenarios

| Escenario | Transacciones/día | Ticket promedio | Ventas/mes (operación plena) |
|-----------|-------------------|-----------------|------------------------------|
| Conservador | 30 | $18,000 | $15.1M |
| Moderado | 45 | $20,000 | $25.2M |
| Optimista | 65 | $22,000 | $40.0M |

**Ramp-up:** Se modeló una curva de arranque realista (40% → 100% en 6 meses), reflejando la entrada gradual al mercado.

---

## 3. Propuesta al Inversionista

### 3.1 Contexto de la Negociación

**Lo que aporta el inversionista:**
- $40,000,000 COP en capital

**Lo que aporta la fundadora:**
- Local propio (ahorro de $3-5M COP/mes en arriendo = $36-60M/año)
- Trabajo diario como operadora (equivalente a $2.5-3.5M/mes en salario de gerente = $30-42M/año)
- Concepto del negocio, plan de negocios, ejecución
- Título profesional en Administración de Negocios (EAFIT)
- Experiencia en retail

**Implicación:** El aporte de la fundadora en año 1 (local + trabajo) equivale a $66-102M COP, significativamente más que los $40M del inversionista. Esto justifica que la fundadora mantenga participación mayoritaria.

### 3.2 Tres Estructuras Propuestas

Se modelaron tres opciones para darle flexibilidad a la negociación:

---

#### OPCIÓN A — Equity Preferente con Recompra (RECOMENDADA)

**Estructura:**
1. El inversionista recibe **30% del equity** de la SAS
2. **Retorno preferente:** Recibe 70% de las utilidades netas distribuibles hasta acumular **$52M COP** ($40M capital + 30% de premium = $12M)
3. **Reducción automática:** Al alcanzar los $52M, su participación baja a **10%**
4. **Opción de recompra:** La fundadora puede comprar el 10% restante a **valor en libros × 1.2** en cualquier momento dentro de los 24 meses siguientes
5. Durante la fase preferente, la fundadora recibe 30% de utilidades

**Proyección (escenario moderado):**
- Recuperación completa ($52M): ~mes 14-16
- TIR para el inversionista: atractiva
- Después de la recuperación, la fundadora tiene 90% del negocio

**Por qué funciona para ambos:**

| Para el inversionista | Para la fundadora |
|----------------------|-------------------|
| Recupera capital + 30% premium | Mantiene 70% desde el inicio |
| Prioridad en distribuciones | No tiene cuota fija mensual |
| Mantiene 10% como upside | Recupera 90-100% del negocio |
| Riesgo moderado | Paga solo cuando hay utilidad |

---

#### OPCIÓN B — Préstamo Privado (Sin Equity)

**Estructura:**
1. Préstamo de $40M COP, **sin participación accionaria**
2. Tasa: **2% mensual** (24% EA — competitiva para préstamo privado en Colombia)
3. **Gracia:** 3 meses (solo intereses: $800K/mes)
4. **Plazo:** 18 meses
5. **Cuota fija (mes 4-18):** ~$3,112,000 COP/mes
6. Total pagado: ~$49.1M COP

**Ventaja:** Sin cesión de equity. Salida limpia a los 18 meses.
**Desventaja:** Cuotas fijas que presionan el flujo de caja, especialmente en meses de ramp-up.

---

#### OPCIÓN C — Revenue Share (Sin Equity)

**Estructura:**
1. **Sin participación accionaria**
2. El inversionista recibe **12% de las ventas brutas** hasta acumular **$56M COP** ($40M × 1.4)
3. Una vez completado, cesa toda obligación

**Proyección (moderado):** Recuperación en ~mes 17-18.
**Ventaja:** Muy simple legalmente, sin equity.
**Desventaja:** Más costoso (40% de premium), se paga sobre ventas (no utilidad), puede presionar caja en meses de bajo margen.

### 3.3 Comparativo Rápido

| Aspecto | Opción A | Opción B | Opción C |
|---------|----------|----------|----------|
| Equity cedida | 30% → 10% → 0% | 0% | 0% |
| Total pagado | $52M (1.3x) | $49.1M (1.23x) | $56M (1.4x) |
| Riesgo para fundadora | Bajo | Medio (cuotas fijas) | Medio |
| Flexibilidad de caja | Alta | Baja | Media |
| Complejidad legal | Media | Baja | Baja |
| Salida del inversionista | Gradual (14-18 meses) | Fija (18 meses) | Gradual (~18 meses) |

### 3.4 Recomendación

**Se recomienda la Opción A** por:

1. **Alineación de incentivos:** El inversionista gana cuando el negocio gana. Esto incentiva su apoyo (contactos, mentoría, promoción del negocio).

2. **Protección del flujo de caja:** En meses difíciles no hay cuota fija. Solo se distribuye cuando hay utilidad real.

3. **Camino claro a la independencia:** La fundadora puede recuperar el 100% del negocio en ~2-3 años sin presión financiera insostenible.

4. **Retorno atractivo para el inversionista:** 30% de premium sobre capital, con exposición al upside del negocio.

5. **Estructura común en emprendimiento:** Es similar al modelo de "preferred equity" usado en venture capital, adaptado a la escala local.

---

## 4. Supuestos Clave y Limitaciones

### Supuestos que más impactan los resultados:

1. **COGS al 32%:** Si el costo de la pulpa de açaí sube, el margen se comprime. El análisis de sensibilidad muestra que con COGS al 40%, la utilidad neta cae significativamente pero el proyecto sigue siendo viable.

2. **Ramp-up de 6 meses:** Se asume que el negocio llega al 100% de capacidad en el mes 7. Si la adopción es más lenta, el payback se extiende.

3. **Sin arriendo:** Esta es la mayor ventaja estructural. Si por alguna razón se pierde el local, toda la evaluación cambia fundamentalmente.

4. **Fundadora como operadora:** El modelo incluye remuneración modesta ($2-3M/mes) pero no un salario de mercado de gerente ($3.5M+). Esto es realista para el arranque pero debe aumentar con el tiempo.

5. **Valor terminal:** Usa un múltiplo de 3.5x EBITDA. Si el negocio no logra estabilizarse, el valor terminal sería menor.

### Limitaciones del modelo:

- No incluye inflación explícita (precios y costos crecen proporcionalmente)
- No modela estacionalidad (Barrancabermeja tiene clima constante, minimizando este efecto)
- No incluye inversión en frozen yogurt (fase 2)
- No incluye costo de oportunidad del local (podría alquilarse a terceros)
- No modela segundo punto de venta

---

## 5. Próximos Pasos

1. **Validar costos reales** de insumos (cotizar açaí, granola, packaging con proveedores)
2. **Definir estructura con inversionista** usando este análisis como base
3. **Formalizar acuerdo** en estatutos de la SAS (cláusulas de preferencia, recompra, anti-dilución)
4. **Consultar con abogado** las cláusulas específicas del acuerdo de accionistas
5. **Actualizar el modelo** una vez se tengan costos reales y se defina la estructura

---

## 6. Archivos del Modelo

| Archivo | Descripción |
|---------|-------------|
| `Bondi_Acai_Evaluacion_Financiera.xlsx` | Modelo financiero completo (10 hojas) |
| `generar_modelo.py` | Script Python generador (reproducible, modificable) |
| `Metodologia_y_Propuesta_Inversionista.md` | Este documento |

Para modificar supuestos y regenerar el Excel, editar las constantes al inicio de `generar_modelo.py` y ejecutar:

```bash
python generar_modelo.py
```
