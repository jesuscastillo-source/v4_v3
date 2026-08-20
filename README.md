# Generador de Documentos RRHH — v3 (descargas Word/PDF separadas)

⚠️ **Repo aparte, en desarrollo.** Construido sobre v2 (cálculo de finiquito
automatizado + PDF). No reemplaces tu app en producción con esta hasta
probarla tú mismo.

## Qué cambia respecto a v2

- Cada pestaña que genera documentos ahora ofrece **dos botones separados**:
  "⬇️ Descargar en Word" y "⬇️ Descargar en PDF" (en vez de un solo ZIP
  mezclado). En la pestaña fusionada de Finiquitos hay un tercero para el
  Excel con los montos.
- El botón de PDF solo aparece si activaste el checkbox de generar PDF y al
  menos uno se generó correctamente.

## Bug que se encontró y arregló en el camino

Streamlit vuelve a ejecutar todo el script cada vez que se aprieta un botón
de descarga. Con un solo botón esto no se notaba, pero con dos botones
separados, al apretar el primero (Word) **el segundo (PDF) desaparecía de
la pantalla** — porque los resultados vivían solo dentro del bloque del
botón "Generar", que se resetea en cada re-ejecución.

Se arregló guardando los resultados en `st.session_state` y mostrando los
botones de descarga **fuera** de ese bloque, para que ambos permanezcan
visibles y funcionales sin importar cuántas veces se descargue uno u otro.
Se probó en las 5 pestañas con navegador real: ambos botones (o los 3, en
la pestaña fusionada) se pueden apretar en cualquier orden, las veces que
se quiera, sin que ninguno desaparezca.

## Todo lo demás
Igual que v2: cálculo de finiquito validado (renuncia voluntaria siempre,
$0 si trabajó menos de 30 días), causal escrita en texto exacto del
desplegable, PDF con LibreOffice preservando formato/negrita real, y
resiliencia fila por fila ante datos inválidos.

---

