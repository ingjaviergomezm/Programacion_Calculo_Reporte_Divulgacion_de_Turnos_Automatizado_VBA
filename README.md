<div align="center">

  <h1>📊 Programación de Turnos (12h) y Análisis de Recargos — Excel + VBA (Colombia)</h1>

  <p>
    Herramienta en <strong>Microsoft Excel + VBA</strong> para <strong>programación mensual de turnos de 12 horas</strong>,
    <strong>cálculo automático de recargos de nómina</strong>, <strong>análisis comparativo por trabajador</strong>,
    <strong>generación de reportes ejecutivos en PDF</strong> y <strong>preparación segura de archivos para distribución</strong>.
  </p>

  <p>
    Alineada con la <strong>legislación laboral colombiana vigente al 13 de enero de 2026</strong>.
    Diseñada como herramienta de <strong>control operativo</strong> y <strong>soporte administrativo</strong>
    (no reemplaza un sistema oficial de nómina).
  </p>

  <p>
    <strong>Autor:</strong> Javier Gómez M. · Ingeniero Industrial · Energías Renovables y Eficiencia Energética · IA aplicada al análisis de datos
  </p>

  <hr style="width: 100%; opacity: .25;" />

</div>

<h2>🎯 Propósito del proyecto</h2>
<p>Resolver de forma integrada y auditable los siguientes problemas operativos en entornos <strong>24/7</strong>:</p>
<ul>
  <li>Programar turnos de 12 horas de manera clara y consistente.</li>
  <li>Identificar y cuantificar recargos según ventanas horarias reales (no solo por día calendario).</li>
  <li>Analizar la distribución de recargos por trabajador y detectar desbalances.</li>
  <li>Detectar alertas asociadas a trabajo dominical reiterado.</li>
  <li>Generar reportes ejecutivos en PDF para <strong>Recursos Humanos</strong>.</li>
  <li>Preparar archivos seguros para enviar al personal sin exponer cálculo interno ni reportes.</li>
  <li>Reducir reprocesos, errores manuales y tiempos de consolidación.</li>
</ul>

<h2>🧩 Contexto operativo: turnos y ventanas horarias</h2>

<h3>Turnos base (12 horas)</h3>
<ul>
  <li><strong>Turno diurno:</strong> 06:00 – 18:00</li>
  <li><strong>Turno nocturno:</strong> 18:00 – 06:00</li>
</ul>
<p>Estas ventanas son la base para la clasificación de eventos y recargos.</p>

<h3>⏱ Ventanas de recargo consideradas (reglas del modelo)</h3>
<p>El modelo traduce la programación de turnos a <strong>eventos sujetos a recargo</strong>, según las siguientes ventanas:</p>

<table>
  <thead>
    <tr>
      <th align="left">Tipo</th>
      <th align="left">Código</th>
      <th align="left">Ventana horaria</th>
      <th align="left">Factor aplicado</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Recargo Nocturno</td>
      <td><strong>RN</strong></td>
      <td>Lunes a sábado: <strong>19:00 – 06:00</strong></td>
      <td><strong>11 h</strong> por evento</td>
    </tr>
    <tr>
      <td>Dominical Diurno</td>
      <td><strong>DD</strong></td>
      <td>Domingo: <strong>06:00 – 18:00</strong></td>
      <td><strong>12 h</strong> por evento</td>
    </tr>
    <tr>
      <td>Trabajo Dominical Nocturno (turno del sábado)</td>
      <td><strong>DN(S)</strong></td>
      <td>Domingo: <strong>00:00 – 06:00</strong></td>
      <td><strong>6 h</strong> por evento</td>
    </tr>
    <tr>
      <td>Dominical Nocturno</td>
      <td><strong>DN</strong></td>
      <td>Domingo: <strong>19:00 – 00:00</strong></td>
      <td><strong>5 h</strong> por evento</td>
    </tr>
    <tr>
      <td>Festivo Diurno</td>
      <td><strong>FD</strong></td>
      <td>Festivo: <strong>06:00 – 18:00</strong></td>
      <td><strong>12 h</strong> por evento</td>
    </tr>
    <tr>
      <td>Festivo Nocturno</td>
      <td><strong>FN</strong></td>
      <td>Festivo: <strong>19:00 – 00:00</strong></td>
      <td><strong>5 h</strong> por evento</td>
    </tr>
  </tbody>
</table>

<p>
  <em>Nota:</em> estas reglas reflejan el diseño actual del libro. Si se reutiliza el proyecto en otro contexto operativo o contractual,
  deben validarse las ventanas y factores.
</p>

<h2>✅ Funcionalidad principal</h2>

<h3>✔ Programación de turnos</h3>
<ul>
  <li>Calendario mensual por trabajador.</li>
  <li>Segmentación por roles operativos.</li>
  <li>Identificación visual de turnos diurnos/nocturnos, domingos, festivos, vacaciones y descansos.</li>
  <li>Diseño enfocado en legibilidad operativa.</li>
</ul>

<h3>✔ Cálculo de recargos de nómina</h3>
<ul>
  <li>Matriz de recargos por trabajador.</li>
  <li>Totales de <strong>eventos</strong> y <strong>horas</strong> por tipo (RN, DD, DN(S), DN, FD, FN).</li>
  <li>Cálculos automáticos a partir de la programación validada.</li>
  <li>Fuente de verdad única para análisis y visualización.</li>
</ul>

<h3>✔ Análisis comparativo (Heatmap)</h3>
<ul>
  <li>Heatmap por trabajador y tipo de recargo.</li>
  <li>Comparación <strong>por columna</strong> (clasificación relativa) para detectar concentración y desbalances.</li>
  <li>Útil para responder: “¿quién concentra más RN?”, “¿quién aparece más en festivos?”</li>
</ul>

<h3>✔ Dashboard ejecutivo (tarjetas + KPI legal)</h3>
<p>
  Bloque analítico integrado en la hoja <strong>Programacion</strong> (filas <strong>39–48</strong>):
</p>
<ul>
  <li><strong>6 tarjetas de resumen:</strong> RN, DD, DN(S), DN, FD, FN (<em>horas</em> y <em>eventos</em>).</li>
  <li><strong>KPI legal:</strong> trabajadores con más de 3 domingos trabajados en el mes (semáforo).</li>
</ul>

<h2>🔁 Actualización automática (sin botones)</h2>
<ul>
  <li>Actualización automática mediante <code>Worksheet_Calculate</code>.</li>
  <li>Cada cambio en la programación recalcula: recargos, heatmap, tarjetas y KPI.</li>
  <li>No requiere acciones manuales para refrescar resultados.</li>
</ul>

<h2>📘 Guía de interpretación</h2>

<h3>📊 Interpretación del Heatmap</h3>
<p>
  El heatmap muestra una matriz donde:
</p>
<ul>
  <li><strong>Filas:</strong> trabajadores</li>
  <li><strong>Columnas:</strong> tipos de recargo (RN, DD, DN(S), DN, FD, FN)</li>
  <li><strong>Valores:</strong> número de eventos en el mes</li>
</ul>

<h4>🧠 Lógica del color (comparación por columna)</h4>
<ul>
  <li>🟢 <strong>Verde:</strong> valores bajos (≤ 33 % del rango de la columna)</li>
  <li>🟡 <strong>Amarillo:</strong> valores medios (33 % – 66 %)</li>
  <li>🔴 <strong>Rojo:</strong> valores altos (≥ 66 %)</li>
</ul>
<p>
  <strong>Importante:</strong> el heatmap no mide carga laboral total. Mide <strong>concentración relativa de recargos por tipo</strong>.
</p>

<h3>📌 Tarjetas KPI (Resumen de Recargos)</h3>
<p>Cada tarjeta muestra:</p>
<pre><code>[Tipo de recargo]
[Total horas] h | [Total eventos] evt</code></pre>

<p>Interpretación:</p>
<ul>
  <li><strong>Eventos:</strong> cuántas veces ocurrió el recargo.</li>
  <li><strong>Horas:</strong> eventos × factor (según la tabla de ventanas).</li>
</ul>

<h3>⚖ KPI legal: trabajo dominical reiterado</h3>
<p><strong>Indicador:</strong> <em>TRABAJADORES &gt; 3 DOMINGOS TRABAJADOS</em></p>
<ul>
  <li>🟢 0 trabajadores → sin alerta</li>
  <li>🟡 1–2 trabajadores → atención</li>
  <li>🔴 3 o más → riesgo elevado</li>
</ul>
<p>
  Este KPI es una <strong>alerta operativa</strong>. No sanciona ni interpreta jurídicamente; apoya revisión operativa y administrativa.
</p>

<h2>📄 Reportes ejecutivos en PDF (RRHH)</h2>
<p>
  El libro incluye macros para generar reportes ejecutivos en PDF destinados a Recursos Humanos, a partir de:
  <strong>una hoja principal (Programacion)</strong> y <strong>una hoja por trabajador</strong> con su formato imprimible.
</p>

<h3>🖨 Flujo de exportación a PDF</h3>
<ol>
  <li>Crear una carpeta local en <code>C:\Users\Public\Documents\</code>.</li>
  <li>Nombre de la carpeta basado en: <code>A1</code> + <code>AA1</code> + <code>AD1</code>.</li>
  <li>Exportar a PDF: hoja <strong>Programacion</strong> y hojas individuales de trabajadores.</li>
  <li>Aplicar un pequeño <em>delay</em> entre exportaciones.</li>
  <li>Notificar finalización y abrir carpeta destino.</li>
</ol>

<h2>✉️ Preparación de archivo para envío al personal (gobierno de la información)</h2>
<p>
  El proyecto incluye un botón <strong>“Preparar archivo para envío”</strong> para distribuir la programación a los trabajadores
  <strong>sin exponer cálculos de nómina, heatmaps, KPIs ni reportes individuales</strong>.
</p>

<h3>🔐 Principio aplicado</h3>
<p><strong>La información sensible no se protege, se excluye del archivo distribuido.</strong></p>

<h3>🧾 Flujo del botón “Preparar archivo para envío”</h3>
<ol>
  <li>Crear una <strong>copia</strong> del libro.</li>
  <li>En la copia:
    <ul>
      <li>Eliminar todas las hojas excepto <strong>Programacion</strong>.</li>
      <li>Limpiar filas <strong>39–48</strong> (recargos, heatmap, KPI y detalles de nómina).</li>
      <li>Convertir todas las fórmulas a valores.</li>
      <li>Guardar como <strong>.xlsx</strong> (sin macros).</li>
    </ul>
  </li>
  <li>Nombrar el archivo:
    <pre><code>Programacion &lt;AA1&gt; &lt;AD1&gt;.xlsx</code></pre>
    Ejemplo:
    <pre><code>Programacion ENERO 2026.xlsx</code></pre>
  </li>
  <li>Intentar abrir el cliente de correo (Outlook) con:
    <ul>
      <li>Para: <code>operadores_cusiana@ocensa.com.co</code></li>
      <li>Adjunto: el archivo generado</li>
      <li>El correo se muestra (<em>Display</em>), no se envía automáticamente</li>
    </ul>
  </li>
  <li>Abrir carpeta destino para ver/adjuntar manualmente si es necesario.</li>
</ol>

<h3>⚠️ Manejo de Error 429 (Outlook no disponible)</h3>
<ul>
  <li>Si Outlook está disponible: se crea el correo con adjunto.</li>
  <li>Si Outlook no está instalado, no configurado o está bloqueado por políticas:
    <ul>
      <li>El archivo se genera correctamente.</li>
      <li>Se notifica al usuario.</li>
      <li>La carpeta se abre para adjuntar manualmente.</li>
    </ul>
  </li>
</ul>

<h2>⚖ Marco legal (Colombia)</h2>
<ul>
  <li>Alineado con legislación laboral colombiana vigente al <strong>13-ene-2026</strong>.</li>
  <li>No liquida salarios ni reemplaza sistemas oficiales.</li>
  <li>Funciona como herramienta de control operativo, auditoría e insumo administrativo.</li>
</ul>

<h2>🔧 Reutilización y adaptación</h2>
<p>Antes de reutilizar el proyecto, validar:</p>
<ul>
  <li>Turnos base (¿12 h?, ¿06–18 / 18–06?).</li>
  <li>Ventanas horarias y factores de recargo.</li>
  <li>Festivos aplicables.</li>
  <li>Umbrales internos/legales (domingos).</li>
  <li>Nombres de hojas y estructura del libro.</li>
  <li>Cliente de correo corporativo (Outlook u otro).</li>
</ul>

<h2>📌 Limitaciones conocidas</h2>
<ul>
  <li>Asume turnos de 12 horas con ventanas definidas.</li>
  <li>Uso controlado de celdas combinadas para el layout.</li>
  <li>Cambios legislativos requieren actualización del modelo.</li>
  <li>Automatización de correo depende del cliente instalado y políticas corporativas.</li>
</ul>

<hr style="opacity: .25;" />

<div align="center">
  <p>
    <strong>Este proyecto está construido para ser entendido, auditado y reutilizado.</strong><br/>
    Documenta reglas, ventanas, KPIs y flujos operativos reales.
  </p>
</div>

