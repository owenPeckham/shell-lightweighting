import numpy as np, igl, time
d=np.load('docs/cshell/research/surface/cant_tess.npz');V,T,F=d['V'],d['T'],d['F']
import pymeshlab as pm
ms=pm.MeshSet(); ms.add_mesh(pm.Mesh(V,T))
ms.meshing_isotropic_explicit_remeshing(targetlen=pm.PureValue(0.75),iterations=10,featuredeg=30,checksurfdist=False)
m=ms.current_mesh();V2=m.vertex_matrix();T2=m.face_matrix()
d3=igl.point_mesh_squared_distance(V,V2,T2)[0]**.5
# vertices on face boundaries
nf=np.zeros(len(V),int); 
sets=[set() for _ in V]
for fi,t in zip(F,T):
    for v in t: sets[v].add(fi)
nf=np.array([len(s) for s in sets])
print('bad total',(d3>0.02).sum(),'of',len(V),'; on boundary:',((d3>0.02)&(nf>1)).sum(),' boundary verts',(nf>1).sum())
# dihedral at boundary
N=igl.per_face_normals(V,T,np.array([1.,0,0])) if False else None
print('max d3 on boundary',d3[nf>1].max(),'max d3 interior',d3[nf==1].max(), 'frac interior bad',(d3[nf==1]>0.02).mean())
print('percentiles interior', np.percentile(d3[nf==1],[50,90,99,99.9]))
