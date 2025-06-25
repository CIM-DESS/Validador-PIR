import io
import json
import pandas as pd
from django.http import HttpResponse
from django.views.decorators.csrf import csrf_exempt
from django.shortcuts import render
from .validacion import procesar_archivos_excel
from django.contrib.auth.decorators import login_required

@login_required
def inicio(request):
    novedades = []
    novedades_json = '[]'
    resumen = None
    error_usuario = None  # ✅ Para mostrar un mensaje claro si hay error de permisos

    if request.method == 'POST':
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
        'error_usuario': error_usuario  # ✅ Lo pasamos a la plantilla
    })

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
    

