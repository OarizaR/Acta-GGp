\documentclass[12pt,letterpaper]{article}

% --------------------------------------------------
% PAQUETES
% --------------------------------------------------
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[spanish]{babel}
\usepackage{geometry}
\usepackage{setspace}
\usepackage{indentfirst} % Sangría en la primera línea del párrafo (APA)
\usepackage{titlesec}    % Personalización de títulos (APA)
\usepackage{fancyhdr}    % Número de página en esquina superior derecha
\usepackage{array}
\usepackage{booktabs}    % Formato de tablas sin líneas verticales (Estilo APA)
\usepackage{longtable}
\usepackage{enumitem}
\usepackage{amsmath}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{hyperref}

% --------------------------------------------------
% CONFIGURACIÓN APA 7ma Edición
% --------------------------------------------------
% Márgenes de 2.54 cm (1 pulgada) en todos los lados
\geometry{
    margin=2.54cm
}

% Doble espacio en todo el documento
\doublespacing

% Sangría de 1.27 cm (0.5 pulgadas) al inicio de cada párrafo
\setlength{\parindent}{1.27cm}

% Configuración del encabezado (Número de página arriba a la derecha)
\pagestyle{fancy}
\fancyhf{}
\renewcommand{\headrulewidth}{0pt}
\fancyhead[R]{\thepage}

% Formato de Títulos estilo APA
% Nivel 1: Centrado, Negrita
\titleformat{\section}[block]{\normalfont\normalsize\bfseries\centering}{\thesection.}{0.5em}{}
% Nivel 2: Alineado a la izquierda, Negrita
\titleformat{\subsection}[block]{\normalfont\normalsize\bfseries}{\thesubsection.}{0.5em}{}
% Nivel 3: Alineado a la izquierda, Negrita, Cursiva
\titleformat{\subsubsection}[block]{\normalfont\normalsize\bfseries\itshape}{\thesubsubsection.}{0.5em}{}

\hypersetup{
    colorlinks=true,
    linkcolor=black,
    urlcolor=blue
}

% --------------------------------------------------
% DOCUMENTO
% --------------------------------------------------
\begin{document}

% Asegurar que la portada tenga el número de página
\thispagestyle{fancy}

% --------------------------------------------------
% PORTADA APA 7ma Edición (Formato Estudiante)
% --------------------------------------------------

\vspace*{3\baselineskip} % 3 a 4 líneas en blanco desde el margen superior

\begin{center}
    \textbf{Acta de Constitución del Proyecto: Optimización de la Gestión de Habitaciones y Precios de un Hotel Mediano Mediante Excel}\\[2\baselineskip]
    
    Tito Samuel Mendez Orjuela, Laura Valentina Guerrero Nivia, Oscar David Ariza Rojas, \\
    Laura Sofía Ruiz Constante y Diego Leonardo Téllez Vargas\\[1\baselineskip]
    
    Facultad de Ingeniería, Universidad Nacional de Colombia\\[1\baselineskip]
    
    Gerencia y Gestión de Proyectos\\[1\baselineskip]
    
    Docente: Juan Sebastian Sanchez\\[1\baselineskip]
    
    16 de septiembre de 2026
\end{center}

\newpage

% --------------------------------------------------
% INTRODUCCIÓN
% --------------------------------------------------

% En APA, el título del documento se repite en la parte superior de la primera página de texto
\section*{Acta de Constitución del Proyecto: Optimización de la Gestión de Habitaciones y Precios de un Hotel Mediano Mediante Excel}

\textbf{Nombre del proyecto:} Optimización de la gestión de habitaciones y precios de un hotel mediano mediante Excel.

\textbf{Problema a solucionar:}

El Hotel Neller Plaza, ubicado en San José del Guaviare, enfrenta tres problemas operativos críticos. En primer lugar, la variabilidad de la demanda genera habitaciones desocupadas en temporada baja y una utilización inadecuada de la capacidad en temporada alta, sin que exista un modelo cuantitativo que oriente la fijación de tarifas. En segundo lugar, la zona presenta cortes frecuentes de energía eléctrica que afectan directamente la experiencia de los huéspedes, la operación de los sistemas informáticos y la conservación de alimentos en refrigeración. En tercer lugar, la gestión administrativa del hotel se realiza de forma manual o con herramientas básicas, lo que genera errores en el registro de reservas, duplicidad de información y tiempos de respuesta elevados en la atención al cliente.

Por estas razones, el proyecto plantea un enfoque integral que combina: (i) un modelo de optimización en Excel para maximizar la utilidad mediante la combinación óptima de tarifas y disponibilidad; (ii) la adquisición e instalación de un sistema de respaldo eléctrico con planta diésel y paneles solares complementarios; y (iii) el diseño de una base de datos SQL y una interfaz gráfica para la gestión eficiente de habitaciones y reservas.

% --------------------------------------------------
% 1.1
% --------------------------------------------------

\section{Propósito y justificación del proyecto}

El proyecto tiene como propósito mejorar de forma integral la operación del Hotel Neller Plaza, ubicado en San José del Guaviare. El hotel cuenta con tres pisos y nueve habitaciones por piso, para un total de 27 habitaciones, con tarifas entre \$110.000 y \$155.000 COP por noche.

Desde el punto de vista analítico, el proyecto desarrollará un modelo de optimización en Excel que permita determinar la combinación de tarifas y disponibilidad que maximice la utilidad del hotel bajo diferentes escenarios de demanda. Sin embargo, la optimización tarifaria pierde efectividad si el hotel no puede garantizar condiciones mínimas de operación. San José del Guaviare es una zona que históricamente ha presentado interrupciones frecuentes en el suministro de energía eléctrica, situación que genera molestias graves en los huéspedes, pérdidas en la cadena de frío y la imposibilidad de operar sistemas informáticos durante los cortes. Por esta razón, el alcance del proyecto incluye la adquisición e instalación de un sistema de respaldo eléctrico compuesto por una planta diésel industrial con capacidad para garantizar al menos 8 horas de autonomía, complementada con un sistema de paneles solares que permita reducir la dependencia del diésel y los costos de operación a largo plazo.

Adicionalmente, el hotel gestiona sus datos de manera precaria: las reservas se registran en cuadernos o en hojas de cálculo aisladas, lo que genera errores, duplicidad de información y demoras en la atención. El proyecto contempla el diseño e implementación de una base de datos relacional (SQL) y una interfaz gráfica (UI) intuitiva que permita a los empleados gestionar habitaciones, reservas y tarifas de forma centralizada, conectada al modelo de optimización desarrollado en Excel.

Desde el punto de vista estratégico y financiero, el proyecto transformará la operación del hotel mediante la integración de tecnología, infraestructura de respaldo y herramientas analíticas de apoyo a la toma de decisiones.

% --------------------------------------------------
% 1.2
% --------------------------------------------------

\section{Descripción y entregables}

El proyecto abarca tres ejes: (a) la construcción de un modelo de optimización en Excel para tarifas y disponibilidad, (b) la instalación de infraestructura de respaldo eléctrico, y (c) el desarrollo de un sistema de información para la gestión de habitaciones.

\subsubsection*{Entregables}

% El uso del paquete booktabs (\toprule, \midrule, \bottomrule) ya cumple con el formato APA para tablas (sin líneas verticales)
\begin{longtable}{>{\centering\arraybackslash}p{0.08\textwidth}
                    >{\raggedright\arraybackslash}p{0.20\textwidth}
                    >{\raggedright\arraybackslash}p{0.34\textwidth}
                    >{\raggedright\arraybackslash}p{0.28\textwidth}}
\toprule
\textbf{Cód.} & \textbf{Entregable} & \textbf{Descripción} & \textbf{Condición de aceptación} \\
\midrule
\endfirsthead
\toprule
\textbf{Cód.} & \textbf{Entregable} & \textbf{Descripción} & \textbf{Condición de aceptación} \\
\midrule
\endhead

E1 & Diagnóstico situacional & Documento con la descripción del hotel, su capacidad, tarifas, demanda, estado de la red eléctrica y del sistema de información actual. & Contiene la información verificable del hotel y define claramente los tres problemas abordados. \\
\midrule
E2 & Base de datos del modelo & Tabla estructurada con datos de habitaciones, tarifas, demanda, costos operativos y costos energéticos. & Los datos son utilizables como entradas del modelo de Excel. \\
\midrule
E3 & Modelo matemático & Formulación de la función objetivo, variables de decisión y restricciones del problema de optimización. & El modelo es coherente y permite la maximización de la utilidad. \\
\midrule
E4 & Modelo en Excel & Archivo Excel con datos, cálculos, función objetivo, restricciones y Solver configurado. & Solver obtiene una solución factible para cada escenario planteado. \\
\midrule
E5 & Análisis de escenarios & Evaluación bajo escenarios de baja, media y alta demanda. & Cada escenario presenta resultados comparables. \\
\midrule
E6 & Sistema de respaldo eléctrico & Planta diésel industrial (20~kW, 8~h autonomía) instalada con transferencia automática, y sistema fotovoltaico complementario (15~kWp). & Pruebas de arranque y autonomía superadas satisfactoriamente. \\
\midrule
E7 & Sistema de información & Base de datos SQL y aplicación con interfaz gráfica para gestión de habitaciones, reservas y tarifas. & La aplicación opera correctamente y se conecta al modelo de optimización. \\
\midrule
E8 & Informe final & Documento que consolide diagnóstico, metodología, modelo, infraestructura, software, resultados y conclusiones. & Presenta el desarrollo completo del proyecto. \\
\midrule
E9 & Presentación final & Material para la sustentación del proyecto. & Explica el problema, la solución integral y las conclusiones. \\
\bottomrule
\end{longtable}

\subsubsection*{Exclusiones del proyecto}

El proyecto no contempla la construcción de nuevas habitaciones, la implementación de sistemas comerciales tipo PMS (Property Management System) de terceros, ni campañas publicitarias. Las recomendaciones derivadas del modelo serán de carácter analítico.

% --------------------------------------------------
% 1.3
% --------------------------------------------------
\section{Requerimientos de alto nivel del proyecto y del producto}

\subsubsection*{Requerimientos del proyecto}

\begin{longtable}{>{\centering\arraybackslash}p{0.12\textwidth}
                    >{\raggedright\arraybackslash}p{0.78\textwidth}}
\toprule
\textbf{Código} & \textbf{Requerimiento} \\
\midrule
\endfirsthead
\toprule
\textbf{Código} & \textbf{Requerimiento} \\
\midrule
\endhead

RP-01 & Recopilar información sobre capacidad, tipos de habitación, tarifas, demanda, costos operativos y costos energéticos del hotel. \\
\midrule
RP-02 & Formular matemáticamente el problema de optimización de tarifas y disponibilidad. \\
\midrule
RP-03 & Desarrollar el modelo de optimización en Excel con Solver. \\
\midrule
RP-04 & Analizar como mínimo tres escenarios de demanda: baja, media y alta. \\
\midrule
RP-05 & Diseñar e instalar el sistema de respaldo eléctrico (planta diésel + paneles solares). \\
\midrule
RP-06 & Diseñar e implementar la base de datos SQL y la interfaz gráfica de gestión. \\
\midrule
RP-07 & Documentar la metodología, resultados y conclusiones del proyecto. \\
\midrule
RP-08 & Entregar todos los productos dentro del plazo establecido. \\
\bottomrule
\end{longtable}

\subsubsection*{Requerimientos del producto}

\begin{longtable}{>{\centering\arraybackslash}p{0.12\textwidth}
                    >{\raggedright\arraybackslash}p{0.78\textwidth}}
\toprule
\textbf{Código} & \textbf{Requerimiento} \\
\midrule
\endfirsthead
\toprule
\textbf{Código} & \textbf{Requerimiento} \\
\midrule
\endhead

RF-01 & El modelo en Excel debe permitir ingresar y modificar datos de habitaciones, tarifas y demanda. \\
\midrule
RF-02 & El modelo debe calcular ingresos, costos (incluyendo costos energéticos) y utilidad. \\
\midrule
RF-03 & Solver debe determinar la combinación óptima de tarifas y disponibilidad. \\
\midrule
RF-04 & El sistema de respaldo debe garantizar al menos 8 horas de autonomía eléctrica. \\
\midrule
RF-05 & El sistema fotovoltaico debe aportar al menos el 30\% de la demanda diurna del hotel. \\
\midrule
RF-06 & La interfaz gráfica debe permitir registrar, consultar y modificar reservas en tiempo real. \\
\midrule
RF-07 & La base de datos SQL debe almacenar el histórico de ocupación, tarifas y huéspedes. \\
\midrule
RF-08 & El sistema de información debe conectarse al modelo de Excel para alimentar las variables de entrada. \\
\bottomrule
\end{longtable}

\subsubsection*{Requerimientos no funcionales}

\begin{longtable}{>{\centering\arraybackslash}p{0.12\textwidth}
                    >{\raggedright\arraybackslash}p{0.78\textwidth}}
\toprule
\textbf{Código} & \textbf{Requerimiento} \\
\midrule
\endfirsthead
\toprule
\textbf{Código} & \textbf{Requerimiento} \\
\midrule
\endhead

RNF-01 & El modelo en Excel debe presentar una estructura organizada con datos de entrada, cálculos y resultados claramente diferenciados. \\
\midrule
RNF-02 & La interfaz gráfica debe ser intuitiva y operable por personal sin formación técnica avanzada. \\
\midrule
RNF-03 & El sistema de información debe permitir actualizar datos sin reconstruir la estructura principal. \\
\midrule
RNF-04 & La planta eléctrica debe arrancar automáticamente mediante transferencia (ATS) en menos de 15 segundos tras un corte. \\
\midrule
RNF-05 & Los resultados del modelo deben ser trazables desde los datos de entrada hasta la solución óptima. \\
\bottomrule
\end{longtable}

% --------------------------------------------------
% 1.4
% --------------------------------------------------

\section{Objetivos medibles y criterios de éxito}

El proyecto busca optimizar la utilidad del Hotel Neller Plaza, garantizar la continuidad eléctrica y modernizar la gestión de información. Los criterios de éxito se presentan a continuación.

\begin{longtable}{>{\raggedright\arraybackslash}p{0.14\textwidth}
                    >{\raggedright\arraybackslash}p{0.38\textwidth}
                    >{\raggedright\arraybackslash}p{0.38\textwidth}}
\toprule
\textbf{Dimensión} & \textbf{Objetivo medible} & \textbf{Criterio de éxito} \\
\midrule
\endfirsthead
\toprule
\textbf{Dimensión} & \textbf{Objetivo medible} & \textbf{Criterio de éxito} \\
\midrule
\endhead

Alcance & Desarrollar modelo de optimización con tarifas, disponibilidad, demanda, ingresos, costos y utilidad. & El modelo integra todas las variables y ejecuta correctamente la optimización. \\
\midrule
Alcance & Analizar mínimo tres escenarios de demanda. & Se obtienen resultados diferenciados y comparables. \\
\midrule
Infraestructura & Instalar planta diésel con 8~h de autonomía y sistema fotovoltaico complementario. & Las pruebas de arranque y autonomía son satisfactorias. \\
\midrule
Software & Implementar base de datos SQL e interfaz gráfica funcional. & La aplicación opera sin errores y se conecta al modelo Excel. \\
\midrule
Calidad & Garantizar coherencia entre datos de entrada y resultados del modelo. & Los resultados son verificables y sin inconsistencias. \\
\midrule
Tiempo & Completar el proyecto dentro del periodo establecido. & Todos los entregables se entregan antes de la fecha límite. \\
\midrule
Costo & Ejecutar el proyecto dentro del presupuesto de \$130.000.000 COP. & El gasto total no supera el presupuesto aprobado. \\
\bottomrule
\end{longtable}

% --------------------------------------------------
% 1.5
% --------------------------------------------------

\section{Supuestos y restricciones}

\subsection{Supuestos}

\begin{itemize}
    \item La demanda estimada representa razonablemente el comportamiento real del hotel.
    \item Los costos operativos permanecen constantes durante el período de análisis.
    \item El hotel opera todos los días del período estudiado.
    \item La radiación solar en San José del Guaviare es suficiente para justificar la inversión fotovoltaica (promedio $\geq$ 4,5 kWh/m$^2$/día).
    \item El hotel cuenta con espacio disponible en la cubierta para la instalación de paneles solares.
    \item Los empleados del hotel recibirán capacitación para operar la nueva interfaz de gestión.
\end{itemize}

\subsection{Restricciones}

\begin{itemize}
    \item El número de habitaciones disponibles es fijo (27 unidades).
    \item Los precios deben mantenerse dentro de rangos competitivos para la zona.
    \item El proyecto debe desarrollarse utilizando Excel como herramienta de optimización.
    \item El presupuesto total aprobado es de \$130.000.000 COP.
    \item El acceso logístico a San José del Guaviare puede generar sobrecostos en el transporte de equipos.
    \item La información disponible puede corresponder a datos estimados por la gerencia.
\end{itemize}

% --------------------------------------------------
% 1.6
% --------------------------------------------------

\section{Riesgos iniciales de alto nivel}

El proyecto integra componentes de software, hardware e infraestructura eléctrica, lo que amplía el espectro de incertidumbres. A continuación se identifican los riesgos principales con su estrategia de respuesta.

\begin{center}
\small
\begin{tabular}{p{0.8cm} p{5.2cm} p{1.3cm} p{5.5cm}}
\toprule
\textbf{Cód.} & \textbf{Riesgo (Causa $\rightarrow$ Efecto)} & \textbf{Nivel} & \textbf{Respuesta preliminar} \\
\midrule
R-01 & Falta de registros históricos: datos insuficientes para calibrar el modelo de demanda. & Alto & Recolectar estimaciones directas de la gerencia y usar distribuciones teóricas. \\
\addlinespace
R-02 & Formulación no lineal que impida a Solver encontrar óptimo global. & Alto & Mantener restricciones lineales y validar con el método Simplex. \\
\addlinespace
R-03 & Demoras en el transporte de la planta diésel y paneles solares hasta San José del Guaviare. & Alto & Contratar proveedor con experiencia en logística a zonas apartadas y establecer holguras en el cronograma. \\
\addlinespace
R-04 & Condiciones climáticas adversas (alta nubosidad, lluvia prolongada) que reduzcan la eficiencia del sistema fotovoltaico. & Medio & Dimensionar los paneles con un factor de seguridad del 20\% y mantener la planta diésel como respaldo primario. \\
\addlinespace
R-05 & Resistencia del personal del hotel a adoptar el nuevo sistema de información. & Medio & Diseñar una interfaz intuitiva y programar sesiones de capacitación práctica. \\
\addlinespace
R-06 & Incompatibilidad entre la red eléctrica local y el sistema de transferencia automática. & Medio & Realizar estudio previo de la acometida eléctrica y contratar un ingeniero electricista certificado. \\
\bottomrule
\end{tabular}
\end{center}

% --------------------------------------------------
% 1.7
% --------------------------------------------------

\section{Cronograma de hitos principales}

Las actividades se estructuran en entregas clave a lo largo de diez semanas, integrando los tres ejes del proyecto.

\begin{center}
\small
\begin{tabular}{p{7cm} p{4cm} p{3.5cm}}
\toprule
\textbf{Hito} & \textbf{Fecha prevista} & \textbf{Evidencia} \\
\midrule
Aprobación del acta de constitución & 16 sept.\ 2026 & Acta firmada \\
Diagnóstico integral y recolección de datos & 23 sept.\ 2026 & Informe diagnóstico \\
Diseño del modelo matemático & 30 sept.\ 2026 & Ecuaciones y restricciones \\
Cotización y compra de equipos (planta, paneles, servidor) & 7 oct.\ 2026 & Órdenes de compra \\
Implementación del modelo en Excel & 14 oct.\ 2026 & Archivo .xlsx funcional \\
Instalación de planta diésel y paneles solares & 28 oct.\ 2026 & Acta de pruebas eléctricas \\
Desarrollo de base de datos SQL e interfaz gráfica & 4 nov.\ 2026 & Aplicación desplegada \\
Validación de escenarios y análisis de resultados & 11 nov.\ 2026 & Análisis de sensibilidad \\
Elaboración del informe final & 18 nov.\ 2026 & Documento borrador \\
Presentación y cierre del proyecto & 25 nov.\ 2026 & Informe y diapositivas \\
\bottomrule
\end{tabular}
\end{center}

% --------------------------------------------------
% 1.8
% --------------------------------------------------

\section{Presupuesto estimado}

El presupuesto aprobado para el proyecto es de \textbf{\$130.000.000 COP} (ciento treinta millones de pesos colombianos). Este monto cubre los tres ejes del proyecto: infraestructura eléctrica, desarrollo tecnológico y gestión analítica. Los precios fueron estimados con base en cotizaciones del mercado colombiano vigentes a septiembre de 2026.

\begin{center}
\small
\begin{tabular}{p{7.5cm} r}
\toprule
\textbf{Rubro} & \textbf{Costo estimado (COP)} \\
\midrule
\multicolumn{2}{l}{\textit{Infraestructura eléctrica}} \\
Planta diésel industrial 20~kW insonorizada + ATS & \$38.000.000 \\
Tanque de combustible (500~gal) + instalación & \$4.500.000 \\
Sistema fotovoltaico 15~kWp (paneles + inversores + estructura) & \$60.000.000 \\
Cableado, protecciones y acometida eléctrica & \$3.500.000 \\
Mano de obra instalación eléctrica (ing.\ certificado) & \$5.000.000 \\
\midrule
\multicolumn{2}{l}{\textit{Infraestructura tecnológica}} \\
Servidor Dell PowerEdge T360 (torre, para SQL y aplicación) & \$8.500.000 \\
UPS para servidor (1~kVA) & \$1.200.000 \\
Computadores de escritorio para recepción (2 unidades) & \$4.800.000 \\
Red local (switch, cableado estructurado, access point) & \$2.000.000 \\
\midrule
\multicolumn{2}{l}{\textit{Desarrollo de software}} \\
Diseño e implementación de base de datos SQL & \$3.500.000 \\
Desarrollo de interfaz gráfica (UI) para gestión de reservas & \$8.000.000 \\
Integración del sistema con el modelo de Excel & \$2.500.000 \\
Pruebas, ajustes y capacitación al personal & \$2.000.000 \\
\midrule
\multicolumn{2}{l}{\textit{Gestión del proyecto}} \\
Transporte de equipos a San José del Guaviare & \$3.500.000 \\
Viáticos del equipo de instalación (5 personas, 7 días) & \$3.500.000 \\
Papelería, impresiones y materiales de presentación & \$500.000 \\
Internet y comunicaciones & \$500.000 \\
Contingencia (5\%) & \$6.000.000 \\
\midrule
\textbf{TOTAL} & \textbf{\$127.500.000} \\
\bottomrule
\end{tabular}
\end{center}

% --------------------------------------------------
% 1.9
% --------------------------------------------------

\section{Lista de interesados (Stakeholders)}

\begin{center}
\small
\begin{tabular}{p{3.8cm} p{6.5cm} p{2cm} p{1.5cm}}
\toprule
\textbf{Interesado} & \textbf{Rol e interés principal} & \textbf{Influencia} & \textbf{Actitud} \\
\midrule
Gerencia Hotel Neller Plaza & Patrocinador. Busca maximizar ingresos, garantizar continuity eléctrica y modernizar la gestión. & Alta & A favor \\
\addlinespace
Docente (Juan S.\ Sanchez) & Evaluador académico. Verifica la aplicación de conceptos de gerencia de proyectos. & Alta & Neutral \\
\addlinespace
Equipo del proyecto & Responsables del desarrollo técnico, análisis matemático, instalación e implementación. & Alta & A favor \\
\addlinespace
Personal de recepción & Usuarios finales del sistema de información. Interés en la facilidad de uso y la reducción de carga operativa. & Media & Neutral \\
\addlinespace
Proveedor eléctrico & Suministro e instalación de planta diésel y sistema fotovoltaico. Interés comercial. & Media & A favor \\
\addlinespace
Huéspedes del hotel & Beneficiarios indirectos. Esperan servicio continuo y tarifas competitivas. & Baja & Neutral \\
\bottomrule
\end{tabular}
\end{center}

% --------------------------------------------------
% 1.10
% --------------------------------------------------

\section{Requisitos de aprobación del proyecto}

Para considerar el proyecto terminado y aprobado se deberán cumplir los siguientes requisitos:

\begin{itemize}
    \item El modelo de Excel deberá estar completamente funcional y generar soluciones factibles mediante Solver.
    \item La planta diésel deberá superar las pruebas de arranque automático y garantizar 8~h de autonomía.
    \item El sistema fotovoltaico deberá estar instalado, conectado y operativo.
    \item La base de datos SQL y la interfaz gráfica deberán operar sin errores críticos.
    \item Deberán analizarse al menos tres escenarios de demanda con resultados coherentes.
    \item El informe final deberá documentar la metodología, resultados e infraestructura instalada.
    \item El proyecto deberá ser presentado y aprobado por el docente.
\end{itemize}

% --------------------------------------------------
% 1.11
% --------------------------------------------------

\section{Asignación del gerente del proyecto y nivel de autoridad}

\textbf{Gerente del proyecto:} Tito Samuel Mendez Orjuela

El gerente del proyecto será responsable de coordinar las actividades del equipo, realizar el seguimiento del cronograma, distribuir las tareas y consolidar los entregables.

Su nivel de autoridad comprenderá:

\begin{itemize}
    \item Asignación y seguimiento de tareas.
    \item Coordinación de reuniones.
    \item Seguimiento del cronograma.
    \item Coordinación de los entregables.
    \item Comunicación con el docente.
    \item Gestión de cambios internos del proyecto.
\end{itemize}

% --------------------------------------------------
% 1.12
% --------------------------------------------------

\section{Personal y recursos preasignados}

\subsection{Personal}

\begin{center}
\begin{tabular}{p{6.5cm} p{6cm}}
\toprule
\textbf{Integrante} & \textbf{Rol} \\
\midrule
Tito Samuel Mendez Orjuela & Gerente del proyecto \\
Laura Valentina Guerrero Nivia & Analista de datos \\
Oscar David Ariza Rojas & Responsable del modelo matemático \\
Laura Sofía Ruiz Constante & Responsable del modelo en Excel \\
Diego Leonardo Téllez Vargas & Documentación y presentación \\
\bottomrule
\end{tabular}
\end{center}

\subsection{Recursos}

\begin{itemize}
    \item Computadores personales y equipos de recepción.
    \item Microsoft Excel con Solver.
    \item Servidor Dell PowerEdge T360 con motor de base de datos SQL.
    \item Planta diésel industrial 20~kW con transferencia automática.
    \item Sistema fotovoltaico de 15~kWp.
    \item Conexión a Internet y red local del hotel.
    \item Herramientas de presentación y documentación.
\end{itemize}

% --------------------------------------------------
% 1.13
% --------------------------------------------------

\section{Aprobaciones}

La aprobación del presente documento establece formalmente el inicio del proyecto y el compromiso de los integrantes con los objetivos, alcance y entregables establecidos.

\vspace{1cm}

\begin{center}
\begin{tabular}{p{6.5cm} p{5cm}}
\toprule
\textbf{Nombre} & \textbf{Firma} \\
\midrule
Tito Samuel Mendez Orjuela & \rule{4cm}{0.4pt} \\[0.8cm]
Laura Valentina Guerrero Nivia & \rule{4cm}{0.4pt} \\[0.8cm]
Oscar David Ariza Rojas & \rule{4cm}{0.4pt} \\[0.8cm]
Laura Sofía Ruiz Constante & \rule{4cm}{0.4pt} \\[0.8cm]
Diego Leonardo Téllez Vargas & \rule{4cm}{0.4pt} \\[0.8cm]
Juan Sebastian Sanchez (Docente) & \rule{4cm}{0.4pt} \\
\bottomrule
\end{tabular}
\end{center}

\vfill

\begin{center}
\textbf{Fecha de aprobación:} 16/09/2026
\end{center}

\newpage
% --------------------------------------------------
% CONTROL DE VERSIONES
% --------------------------------------------------

\section*{Control de versiones}

\begin{center}
\begin{tabular}{p{1.5cm} p{2.5cm} p{6cm} p{4cm}}
\toprule
\textbf{Versión} & \textbf{Fecha} & \textbf{Descripción} & \textbf{Responsable} \\
\midrule
0.1 & 05/09/2026 & Borrador inicial & Equipo del proyecto \\
0.2 & 10/09/2026 & Ajustes posteriores a revisión & Gerente del proyecto \\
1.0 & 16/09/2026 & Versión presentada para aprobación & Gerente del proyecto \\
\bottomrule
\end{tabular}
\end{center}

\newpage
% --------------------------------------------------
% REFERENCIAS
% --------------------------------------------------

\section*{Referencias}

% La sangría francesa (hanging indent) es obligatoria en las referencias de APA
\begin{flushleft}
\setlength{\parindent}{-1.27cm}
\setlength{\leftskip}{1.27cm}

Project Management Institute. (2021). \textit{A guide to the project management body of knowledge (PMBOK\textregistered{} guide) and the standard for project management} (7th ed.). Project Management Institute.

Project Management Institute. (2022). \textit{Process groups: A practice guide}. Project Management Institute.

\end{flushleft}

\end{document}
