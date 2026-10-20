# **Documentación oficial: cálculo de OEE de la línea de embotellado**

Planta embotelladora de agua · Proyecto de cálculo de OEE en Excel con VBA

* **Versión del documento:** v0.1  
* **Última actualización:** \[fecha\]

## **0\. Reglas de este documento**

* Este documento es la **única fuente de verdad** del proyecto.

## **1\. Objetivo y alcance**

### **1.1 Objetivo**

Calcular el OEE (Eficiencia General de los Equipos) de \[nombre de la línea\] mediante un libro de Excel con VBA, con reportes por turno, día y semana.

### **1.2 Alcance**

* Línea medida: \[Esperando definición\]  
* Equipo o punto de medición (por ejemplo, llenadora o fin de línea): \[Esperando definición\]  
* Producto(s) envasado(s): \[Esperando definición\]

### **1.3 Fuera de alcance (por ahora)**

* Sistema completo para otras áreas de la planta.  
* Base de datos real y aplicación web.  
* Otras líneas de producción.

### **1.4 Criterio de éxito**

El cálculo del libro coincide con el turno real calculado a mano (sección 7), con la tolerancia que apruebe el ingeniero: \[Esperando definición\].

## **2\. Equipo y roles**

* **Ingeniero (dueño del proyecto):** define reglas y fórmulas, aprueba cambios.  
* **Joaquín (desarrollo):** implementa, documenta y propone.  
* **Tercera persona:** \[rol por definir\].  
* **Responsable del archivo maestro de Excel:** \[Esperando definición\]  
* **Contacto para dudas rápidas:** \[nombre y medio\]

## **3\. Glosario y fórmulas**

### **3.1 Fórmulas generales del OEE**

Estado: a confirmar con el ingeniero.

* **OEE \= Disponibilidad × Rendimiento × Calidad**  
* **Disponibilidad** \= tiempo operativo / tiempo planificado de producción  
* **Rendimiento** \= (unidades producidas × tiempo de ciclo ideal) / tiempo operativo  
* **Calidad** \= unidades buenas / unidades totales  
* **Unidades buenas** \= unidades totales − unidades rechazadas

### **3.2 Escalera de tiempos**

1. Tiempo calendario.  
2. Tiempo planificado de producción (se descuenta lo que no se planeó producir).  
3. Tiempo operativo (se descuentan las paradas).  
4. Tiempo plenamente productivo (se descuentan pérdidas de velocidad y rechazos).

### **3.3 Las tres familias de pérdida**

* **Disponibilidad:** averías, cambios de formato, ajustes.  
* **Rendimiento:** microparadas y velocidad reducida.  
* **Calidad:** rechazos y producto defectuoso.

### **3.4 Glosario**

* **Tiempo planificado de producción:** \[definición de la planta\]  
* **Tiempo operativo:** \[definición de la planta\]  
* **Parada planificada:** \[definición de la planta\]  
* **Parada no planificada:** \[definición de la planta\]  
* **Microparada:** \[definición y umbral\]  
* **Tiempo de ciclo ideal / velocidad nominal:** \[definición y origen\]  
* **Rechazo:** \[definición de la planta\]  
* **Turno:** \[definición de la planta\]

## **4\. Definiciones particulares de la planta**

Cada punto se completa con el ingeniero. Formato: decisión, fecha, aprobado por, estado.

### **4.1 Línea y equipo medido**

* Decisión: \[ \]  
* Fecha / aprobado por: \[ \]  
* Estado: Esperando definición

### **4.2 Turnos y tiempo planificado**

* ¿Cuántos turnos hay y en qué horarios? ¿Alguno cruza la medianoche?  
* ¿Qué se considera tiempo planificado? ¿Cómo se tratan feriados y días sin producción?  
* Decisión: \[ \]  
* Fecha / aprobado por: \[ \]  
* Estado: Esperando definición

### **4.3 Clasificación de paradas**

Para cada tipo de parada, definir si **descuenta del tiempo planificado** o si **cuenta como pérdida de disponibilidad**:

* Almuerzo / descanso: \[ \]  
* Limpieza y sanitizado: \[ \]  
* Cambio de formato o producto: \[ \]  
* Mantenimiento programado: \[ \]  
* Falta de pedidos o de plan de producción: \[ \]  
* Falta de insumos o materia prima: \[ \]  
* Averías: \[ \]  
* Otras: \[ \]  
* Fecha / aprobado por: \[ \]  
* Estado: Esperando definición

### **4.4 Umbral de registro de paradas**

* ¿Desde cuántos minutos se registra una parada? \[ \]  
* ¿Qué pasa con las paradas más cortas? \[ \]  
* Estado: Esperando definición

### **4.5 Velocidad nominal / tiempo de ciclo ideal**

* Valor por producto o formato: \[ \]  
* Origen del dato (placa de la máquina, histórico, objetivo): \[ \]  
* Quién lo actualiza y cuándo: \[ \]  
* Estado: Esperando definición

### **4.6 Productos y unidad de conteo**

* Productos y formatos (botellas, bidones, otros): \[ \]  
* Unidad de producción (botella, bidón, pack, caja): \[ \]  
* Punto de conteo (salida de llenadora, fin de línea): \[ \]  
* Estado: Esperando definición

### **4.7 Rechazos**

* ¿Qué se considera rechazo? \[ \]  
* ¿El retrabajo o reproceso cuenta como rechazo? \[ \]  
* ¿Quién lo registra, cuándo y en qué punto de la línea? \[ \]  
* Estado: Esperando definición

### **4.8 Cambios de producto dentro de un turno**

* ¿Cómo se reparte el tiempo y la producción entre productos? \[ \]  
* Estado: Esperando definición

### **4.9 Carga de datos**

* ¿Quién carga los datos y cuándo (durante o al final del turno)? \[ \]  
* Detalle de las paradas: hora de inicio y fin, o solo minutos \[ \]  
* Origen de la producción: contador de máquina o anotación manual \[ \]  
* Estado: Esperando definición

### **4.10 Reportes y metas**

* Reportes requeridos (turno, día, semana, otros): \[ \]  
* Meta de OEE y de cada factor: \[ \]  
* Estado: Esperando definición

## **5\. Catálogos**

Se completan a partir de las definiciones de la sección 4\.

### **5.1 Causas de parada**

Por cada causa: código, nombre, categoría (planificada o no planificada) e impacto (descuenta del planificado o pérdida de disponibilidad).

* \[código\] · \[causa\] · \[categoría\] · \[impacto\]

### **5.2 Turnos**

* \[nombre\] · \[hora de inicio\] · \[hora de fin\]

### **5.3 Productos**

* \[código\] · \[producto\] · \[unidad\] · \[velocidad nominal\]

### **5.4 Tipos de rechazo**

* \[código\] · \[tipo de rechazo\] · \[descripción\]

## **6\. Reglas de negocio**

Cada regla tiene un identificador, un estado y una fecha de aprobación.

* **RN-01:** los OEE de períodos más largos (día, semana) se calculan sumando tiempos y unidades de los turnos, nunca promediando los OEE de cada turno. Estado: a confirmar.  
* **RN-02:** unidades buenas \= unidades totales − unidades rechazadas. Estado: a confirmar.  
* **RN-03:** \[Esperando definición\]  
* **RN-04:** \[Esperando definición\]

## **7\. Caso de prueba: turno real calculado a mano**

Este caso es la referencia para validar el cálculo del libro.

### **7.1 Datos de entrada**

* Fecha y turno: \[ \]  
* Horario del turno: \[ \]  
* Tiempo planificado de producción (min): \[ \]  
* Paradas (causa, inicio, fin, minutos): \[ \]  
* Tiempo de ciclo ideal o velocidad nominal: \[ \]  
* Unidades producidas: \[ \]  
* Unidades rechazadas: \[ \]

### **7.2 Cálculo manual**

* Tiempo operativo: \[ \]  
* Disponibilidad: \[ \]  
* Rendimiento: \[ \]  
* Calidad: \[ \]  
* **OEE esperado:** \[ \]

### **7.3 Resultado del libro y diferencia**

* OEE calculado por el libro: \[ \]  
* Diferencia: \[ \]  
* Aprobado por / fecha: \[ \]

### **7.4 Casos adicionales**

* Turno sin producción: \[ \]  
* Turno con cambio de producto: \[ \]  
* Turno con paradas que cruzan medianoche: \[ \]  
* Turno con datos incompletos: \[ \]

## **8\. Estructura de datos prevista**

* **Hoja Paradas:** fecha, turno, línea, causa (lista desplegable), hora de inicio y fin o minutos.  
* **Hoja Producción:** fecha, turno, línea, producto, unidades producidas, unidades rechazadas, velocidad nominal.  
* **Hoja Catálogos:** causas de parada, turnos, productos, tipos de rechazo.  
* **Formularios de carga** con validaciones.  
* **Dashboard:** OEE y sus tres factores, tendencia y desglose de pérdidas.

## **9\. Reportes y dashboard**

* Reporte por turno: \[ \]  
* Reporte por día: \[ \]  
* Reporte por semana: \[ \]  
* Indicadores a mostrar: \[ \]  
* Estado: Esperando definición

## **10\. Plan de entregas**

* **Etapa 1:** cálculo funcionando con datos cargados a mano. Terminado cuando coincide con el caso de prueba de la sección 7\.  
* **Etapa 2:** formularios de carga con validaciones. Terminado cuando: \[ \]  
* **Etapa 3:** dashboard con tendencia y desglose de pérdidas. Terminado cuando: \[ \]  
* Fechas acordadas: \[Esperando definición\]

## **11\. Riesgos y supuestos**

### **Riesgos**

* Definiciones incompletas.  
* Datos desordenados o inconsistentes.  
* Cambios de requisitos durante el desarrollo.  
* Registros sin horarios de paradas.

### **Supuestos**

* Se usa Excel de escritorio con VBA en la notebook del ingeniero.  
* \[otros supuestos acordados\]

## **12\. Preguntas abiertas**

Por cada pregunta: número, pregunta, fecha, responsable, estado.

* Q-01 · \[pregunta\] · \[fecha\] · \[responsable\] · \[abierta\]  
* Q-02 · \[pregunta\] · \[fecha\] · \[responsable\] · \[abierta\]

## **13\. Registro de cambios**

Por cada cambio: fecha, qué cambió, quién lo propuso, quién lo aprobó.

* \[fecha\] · Creación del documento (v0.1) · Joaquín · \[pendiente de aprobación\]

## **14\. Control de versiones del archivo Excel**

* Existe **un único archivo maestro**, con una persona responsable de modificarlo: \[nombre\].  
* Cada versión lleva fecha en el nombre, por ejemplo oee\_v0.3\_AAAA-MM-DD.xlsm.  
* Los demás prueban siempre sobre copias, nunca sobre el maestro.  
* El código (módulos y formularios) se exporta a Git.

