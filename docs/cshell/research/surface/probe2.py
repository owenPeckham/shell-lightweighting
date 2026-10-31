import sys,time,numpy as np, igl
d=np.load('docs/cshell/research/surface/cant_tess.npz');V,T,F=d['V'],d['T'],d['F']
import pymeshlab as pm
def run(**kw):
    ms=pm.MeshSet(); ms.add_mesh(pm.Mesh(V,T))
    t=time.time()
    ms.meshing_isotropic_explicit_remeshing(targetlen=pm.PureValue(0.75),**kw)
    m=ms.current_mesh();V2=m.vertex_matrix();T2=m.face_matrix()
    d1=igl.point_mesh_squared_distance(V2,V,T)[0]**.5
    # sample face centroids & edge midpoints
    C=V2[T2].mean(1); d2=igl.point_mesh_squared_distance(C,V,T)[0]**.5
    d3=igl.point_mesh_squared_distance(V,V2,T2)[0]**.5
    print(kw,m.vertex_number(),'%.1fs'%(time.time()-t),'v->in %.4f cen->in %.4f in->out %.4f'%(d1.max(),d2.max(),d3.max()), 'frac>0.02 %.4f'%(d1>0.02).mean())
run(iterations=10,featuredeg=30,checksurfdist=False)
run(iterations=10,featuredeg=30,checksurfdist=True,maxsurfdist=pm.PureValue(0.02))
run(iterations=10,featuredeg=15,checksurfdist=True,maxsurfdist=pm.PureValue(0.02))
run(iterations=5,featuredeg=30,checksurfdist=True,maxsurfdist=pm.PureValue(0.02),smoothflag=False)
