import io
import json
import pandas as pd
import threading
import uuid
import time
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.shortcuts import render
from .validacion import procesar_archivos_excel
from django.contrib.auth.decorators import login_required

# 🚀 ALMACÉN GLOBAL DE TAREAS ASÍNCRONAS
tareas_procesamiento = {}
tareas_lock = threading.Lock()

def limpiar_tareas_antiguas():
    """Limpia tareas completadas que tengan más de 30 minutos"""
    with tareas_lock:
        ahora = time.time()
        tareas_a_eliminar = []
        
        for tarea_id, tarea in tareas_procesamiento.items():
            if ahora - tarea.get('timestamp', 0) > 1800:  # 30 minutos
                tareas_a_eliminar.append(tarea_id)
        
        for tarea_id in tareas_a_eliminar:
            del tareas_procesamiento[tarea_id]
            
        if tareas_a_eliminar:
            print(f"🧹 Limpiadas {len(tareas_a_eliminar)} tareas antiguas")

def procesar_archivos_async(archivos_data, tarea_id):
    """Procesa archivos en hilo separado"""
    try:
        with tareas_lock:
            tareas_procesamiento[tarea_id]['estado'] = 'procesando'
            tareas_procesamiento[tarea_id]['progreso'] = 10
        
        print(f"🔄 Iniciando procesamiento asíncrono: {tarea_id}")
        
        # Recrear archivos desde datos serializados
        archivos = []
        for archivo_data in archivos_data:
            # Crear un archivo temporal en memoria
            archivo_temp = io.BytesIO(archivo_data['contenido'])
            archivo_temp.name = archivo_data['nombre']
            archivos.append(archivo_temp)
        
        with tareas_lock:
            tareas_procesamiento[tarea_id]['progreso'] = 30
        
        # Procesar archivos
        novedades = procesar_archivos_excel(archivos)
        
        with tareas_lock:
            tareas_procesamiento[tarea_id]['progreso'] = 90
        
        # Calcular resumen
        hojas_con_errores = set((n['archivo'], n['hoja']) for n in novedades)
        resumen = {
            "total_archivos": len(archivos),
            "total_novedades": len(novedades),
            "total_hojas_con_errores": len(hojas_con_errores)
        }
        
        # Marcar como completado
        with tareas_lock:
            tareas_procesamiento[tarea_id] = {
                'estado': 'completado',
                'progreso': 100,
                'novedades': novedades,
                'resumen': resumen,
                'timestamp': time.time()
            }
        
        print(f"✅ Procesamiento completado: {tarea_id}")
        
    except Exception as e:
        print(f"❌ Error en procesamiento asíncrono: {e}")
        with tareas_lock:
            tareas_procesamiento[tarea_id] = {
                'estado': 'error',
                'progreso': 0,
                'error': str(e),
                'timestamp': time.time()
            }

@login_required
def inicio(request):
    novedades = []
    novedades_json = '[]'
    resumen = None
    error_usuario = None

    if request.method == 'POST':
        # Verificar si son resultados de procesamiento asíncrono
        if 'resultados_async' in request.POST:
            try:
                datos_async = json.loads(request.POST.get('resultados_async'))
                novedades = datos_async.get('novedades', [])
                resumen = datos_async.get('resumen', {})
                novedades_json = json.dumps(novedades, ensure_ascii=False)
                
                return render(request, 'interfaz/inicio.html', {
                    'novedades': novedades,
                    'novedades_json': novedades_json,
                    'resumen': resumen,
                    'error_usuario': None
                })
            except Exception as e:
                error_usuario = f"Error procesando resultados asíncronos: {str(e)}"
        
        # Verificar si es una petición para iniciar procesamiento asíncrono
        elif request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return iniciar_procesamiento_async(request)
        
        # Procesamiento síncrono tradicional
        else:
            archivos = request.FILES.getlist('archivos_excel')
            print("Archivos recibidos:", [f.name for f in archivos])

            try:
                if archivos:
                    novedades = procesar_archivos_excel(archivos)
                    novedades_json = json.dumps(novedades, ensure_ascii=False)

                    hojas_con_errores = set((n['archivo'], n['hoja']) for n in novedades)
                    resumen = {
                        "total_archivos": len(archivos),
                        "total_novedades": len(novedades),
                        "total_hojas_con_errores": len(hojas_con_errores)
                    }

            except PermissionError as e:
                print("❌ Error de permisos:", e)
                error_usuario = (
                    "⚠️ El archivo pivote está abierto o en uso. "
                    "Por favor, ciérralo antes de continuar."
                )

            except Exception as e:
                print("❌ Otro error inesperado:", e)
                error_usuario = f"⚠️ Ocurrió un error inesperado: {str(e)}"

    return render(request, 'interfaz/inicio.html', {
        'novedades': novedades,
        'novedades_json': novedades_json,
        'resumen': resumen,
        'error_usuario': error_usuario
    })

def iniciar_procesamiento_async(request):
    """Inicia procesamiento asíncrono y devuelve ID de tarea"""
    try:
        archivos = request.FILES.getlist('archivos_excel')
        
        if not archivos:
            return JsonResponse({'error': 'No se recibieron archivos'}, status=400)
        
        # Limpiar tareas antiguas
        limpiar_tareas_antiguas()
        
        # Crear ID único para la tarea
        tarea_id = str(uuid.uuid4())
        
        # Serializar archivos para pasar al hilo
        archivos_data = []
        for archivo in archivos:
            archivos_data.append({
                'nombre': archivo.name,
                'contenido': archivo.read()
            })
        
        # Crear entrada de tarea
        with tareas_lock:
            tareas_procesamiento[tarea_id] = {
                'estado': 'iniciando',
                'progreso': 0,
                'timestamp': time.time()
            }
        
        # Iniciar procesamiento en hilo separado
        thread = threading.Thread(
            target=procesar_archivos_async,
            args=(archivos_data, tarea_id),
            daemon=True
        )
        thread.start()
        
        return JsonResponse({
            'tarea_id': tarea_id,
            'estado': 'iniciado'
        })
        
    except Exception as e:
        print(f"❌ Error iniciando procesamiento asíncrono: {e}")
        return JsonResponse({'error': str(e)}, status=500)

@csrf_exempt
def estado_procesamiento(request, tarea_id):
    """Devuelve el estado actual de una tarea de procesamiento"""
    try:
        with tareas_lock:
            tarea = tareas_procesamiento.get(tarea_id)
        
        if not tarea:
            return JsonResponse({'error': 'Tarea no encontrada'}, status=404)
        
        # Preparar respuesta
        respuesta = {
            'estado': tarea['estado'],
            'progreso': tarea.get('progreso', 0)
        }
        
        # Si está completado, incluir resultados
        if tarea['estado'] == 'completado':
            respuesta.update({
                'novedades': tarea.get('novedades', []),
                'resumen': tarea.get('resumen', {}),
                'novedades_json': json.dumps(tarea.get('novedades', []), ensure_ascii=False)
            })
        elif tarea['estado'] == 'error':
            respuesta['error'] = tarea.get('error', 'Error desconocido')
        
        return JsonResponse(respuesta)
        
    except Exception as e:
        print(f"❌ Error consultando estado: {e}")
        return JsonResponse({'error': str(e)}, status=500)

@csrf_exempt
def descargar_excel(request):
    if request.method == 'POST':
        data_json = request.POST.get('novedades_json', '[]')
        print("DATA JSON:", data_json)  # 👈 Imprime el contenido recibido

        data = json.loads(json.loads(f'"{data_json}"'))

        df = pd.DataFrame(data)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Novedades')

        output.seek(0)
        response = HttpResponse(
            output.read(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename="novedades.xlsx"'
        return response
    

