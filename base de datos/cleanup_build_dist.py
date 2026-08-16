import os
import shutil

root = os.path.dirname(__file__)
removed = []
for name in ('build', 'dist'):
    path = os.path.join(root, name)
    if os.path.exists(path):
        try:
            shutil.rmtree(path)
            removed.append(path)
        except Exception as e:
            print(f"No se pudo eliminar {path}: {e}")

if removed:
    print('Eliminadas:', removed)
else:
    print('No se encontraron carpetas build/ o dist/ para eliminar.')
