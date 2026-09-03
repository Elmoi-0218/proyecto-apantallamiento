import win32com.client

# Conectar a Inventor
inventor = win32com.client.Dispatch("Inventor.Application")
inventor.Visible = True

# Constantes
kPartDocumentObject = 12291  # Valor numérico de kPartDocumentObject

# Crear un nuevo documento de pieza
part_document = inventor.Documents.Add(kPartDocumentObject, inventor.FileManager.GetTemplateFile(kPartDocumentObject))

# Obtener el componente de definición del documento
comp_def = part_document.ComponentDefinition

# Obtener el gestor de transacciones
trans_mgr = part_document.TransactionManager

# Iniciar una transacción
trans = trans_mgr.StartTransaction(part_document, "Crear pieza")

# Crear un nuevo boceto en el plano X-Y
sketch = comp_def.Sketches.Add(comp_def.WorkPlanes.Item(3))

# Definir los puntos para el perfil del boceto (un rectángulo en este caso)
point1 = sketch.SketchPoints.Add(inventor.TransientGeometry.CreatePoint2d(0, 0))
point2 = sketch.SketchPoints.Add(inventor.TransientGeometry.CreatePoint2d(5, 0))
point3 = sketch.SketchPoints.Add(inventor.TransientGeometry.CreatePoint2d(5, 5))
point4 = sketch.SketchPoints.Add(inventor.TransientGeometry.CreatePoint2d(0, 5))

# Crear líneas entre los puntos para formar un rectángulo
sketch.SketchLines.AddByTwoPoints(point1, point2)
sketch.SketchLines.AddByTwoPoints(point2, point3)
sketch.SketchLines.AddByTwoPoints(point3, point4)
sketch.SketchLines.AddByTwoPoints(point4, point1)

# Crear un perfil para la extrusión
profile = sketch.Profiles.AddForSolid()

# Crear una extrusión de 10 unidades de longitud
extrude_feature = comp_def.Features.ExtrudeFeatures.AddByDistanceExtent(profile, 10, inventor.kPositiveExtentDirection, inventor.kJoinOperation)

# Crear un agujero en la cara superior de la extrusión
face = extrude_feature.Faces.Item(5)
hole_sketch = comp_def.Sketches.Add(face)
hole_center = hole_sketch.SketchPoints.Add(inventor.TransientGeometry.CreatePoint2d(2.5, 2.5))
hole = comp_def.Features.HoleFeatures.AddDrilledByThroughAllExtent(hole_sketch.Profiles.AddForSolid(), inventor.kJoinOperation)

# Aplicar un empalme (fillet) en una de las aristas
edge = extrude_feature.Edges.Item(1)
fillet = comp_def.Features.FilletFeatures.AddSimple([edge], 0.5, True)

# Finalizar la transacción
trans.End()

print("Pieza creada con éxito.")
