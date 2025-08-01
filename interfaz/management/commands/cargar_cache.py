from django.core.management.base import BaseCommand
from interfaz.validacion import inicializar_cache_pivote

class Command(BaseCommand):
    help = 'Pre-carga el cache del archivo pivote para optimizar el rendimiento'

    def handle(self, *args, **options):
        self.stdout.write(
            self.style.SUCCESS('🚀 Iniciando pre-carga del cache del archivo pivote...')
        )
        
        try:
            inicializar_cache_pivote()
            self.stdout.write(
                self.style.SUCCESS('✅ Cache del archivo pivote cargado correctamente.')
            )
            self.stdout.write('💡 El sistema ahora estará más rápido en las validaciones.')
            
        except Exception as e:
            self.stdout.write(
                self.style.ERROR(f'❌ Error cargando cache: {e}')
            )
            raise
