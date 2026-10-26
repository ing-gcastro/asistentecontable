# ADR-0009: Notificación por mail a sectores: conectada pero desactivada

- **Estado:** Aceptado
- **Fecha:** 2026-10-26

## Contexto
Se pidió avisar por mail a cada sector qué facturas tiene para autorizar. Faltan definiciones de negocio (destinatarios, frecuencia, buzón, adjuntos).

## Decisión
`notificador_sectores.py` arma un mail HTML por sector y lo entrega según `NOTIFICACION_MODO`: `desactivado` (default: solo vista previa en `logs\notificaciones`), `borrador` (Borradores de Outlook) o `enviar`. Se llama desde el Paso 2 (aprobadas pendientes) y desde `--pendientes` (recordatorio de todo lo pendiente). Un sector sin destinatarios nunca recibe mails. Las decisiones abiertas están en SDD §8.

## Alternativas consideradas
- SMTP directo: requiere credenciales y configuración del servidor; Outlook ya está autenticado.

## Consecuencias
- Enviar un mail es irreversible: por eso el default no envía.
- Activarlo es solo configuración.
