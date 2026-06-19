# Sistema de Solicitud de Vacaciones

## Descripcion

Proyecto desarrollado para la materia Organizacion Empresarial.

El sistema automatiza el proceso de solicitud de vacaciones mediante un chatbot por consola implementado en Python. Permite registrar solicitudes de vacaciones, validar empleados mediante DNI y gestionar la aprobacion o rechazo de solicitudes.

La informacion se almacena de forma persistente mediante archivos Excel.

## Funcionalidades

* Validacion de empleados por DNI.
* Registro de solicitudes de vacaciones.
* Gestion de estados:
  * PENDIENTE
  * APROBADA
  * RECHAZADA
* Actualizacion automatica del saldo de dias disponibles.
* Persistencia mediante archivos Excel.

## Tecnologias Utilizadas

* Python 3.14
* OpenPyXL
* Excel (.xlsx)
* BPMN 2.0

## Estructura del Proyecto

```text
src/
    vacaciones.py

data/
    empleados.xlsx
    solicitudes.xlsx

docs/
    TPI-OE-SantiagoMeza.pdf
    BPMN_AS_IS.png
    BPMN_TO_BE.png
```

## Instalacion

Instalar la dependencia necesaria:

```bash
pip install openpyxl
```

## Ejecucion

Desde la raiz del proyecto ejecutar:

```bash
python src/vacaciones.py
```

## Archivos de Datos

### empleados.xlsx

| Campo            | Descripcion                  |
| ---------------- | ---------------------------- |
| DNI              | Documento del empleado       |
| Nombre           | Nombre completo              |
| Dias_Disponibles | Dias de vacaciones restantes |

### solicitudes.xlsx

| Campo            | Descripcion                     |
| ---------------- | ------------------------------- |
| ID               | Identificador unico             |
| DNI              | Documento del empleado          |
| Nombre           | Nombre del empleado             |
| Fecha_Inicio     | Inicio de vacaciones            |
| Fecha_Fin        | Fin de vacaciones               |
| Dias_Solicitados | Cantidad de dias                |
| Estado           | PENDIENTE, APROBADA o RECHAZADA |

## Autor

Santiago Meza