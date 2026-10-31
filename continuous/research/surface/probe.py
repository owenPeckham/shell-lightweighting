import sys,time,numpy as np
sys.path.insert(0,'.')
from vshell import geometry as g
t=time.time()
s=g.load_step('examples/step/Solid Cantilever.step')
V,T,F=g.tessellate(s,0.02,0.2)
print(len(V),len(T),time.time()-t, 'faces',len(set(F)))
import igl
print('area',igl.doublearea(V,T).sum()/2, 'bbox',V.min(0),V.max(0))
np.savez('continuous/research/surface/cant_tess.npz',V=V,T=T,F=F)
import pymeshlab as pm
ms=pm.MeshSet(); ms.add_mesh(pm.Mesh(V,T))
print([ (k) for k in ms.filter_parameter_values('meshing_isotropic_explicit_remeshing').keys()] if hasattr(ms,'filter_parameter_values') else '')
t=time.time()
ms.meshing_isotropic_explicit_remeshing(iterations=10,targetlen=pm.PureValue(0.75),featuredeg=30,checksurfdist=True,maxsurfdist=pm.PureValue(0.02))
m=ms.current_mesh(); print('remesh',m.vertex_number(),time.time()-t)
V2=m.vertex_matrix();T2=m.face_matrix()
d1=igl.point_mesh_squared_distance(V2,V,T)[0]**.5
print('one-sided hd out->in',d1.max())
