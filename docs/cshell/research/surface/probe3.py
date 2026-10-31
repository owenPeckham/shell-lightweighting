import sys,time,numpy as np, igl
sys.path.insert(0,'.')
from vshell import geometry as g
s=g.load_step('examples/step/Solid Cantilever.step'); tab=g.face_table(s)
d=np.load('continuous/research/surface/cant_tess.npz');V,T,F=d['V'],d['T'],d['F']
import pymeshlab as pm
def run(**kw):
    ms=pm.MeshSet(); ms.add_mesh(pm.Mesh(V,T))
    t=time.time()
    ms.meshing_isotropic_explicit_remeshing(targetlen=pm.PureValue(0.75),**kw)
    m=ms.current_mesh();V2=m.vertex_matrix();T2=m.face_matrix()
    d1=igl.point_mesh_squared_distance(V2,V,T)[0]**.5
    # dense samples on out mesh: centroids + edge midpoints
    E=np.concatenate([V2[T2[:,[0,1]]].mean(1),V2[T2[:,[1,2]]].mean(1),V2[T2[:,[2,0]]].mean(1),V2[T2].mean(1)])
    d2=igl.point_mesh_squared_distance(E,V,T)[0]**.5
    # in->out on input vertices and tri centroids
    Ci=np.concatenate([V,V[T].mean(1)]); d3=igl.point_mesh_squared_distance(Ci,V2,T2)[0]**.5
    print(kw,m.vertex_number(),'%.1fs'%(time.time()-t),'v->in %.4f samp->in %.4f in->out %.4f'%(d1.max(),d2.max(),d3.max()),'frac(in->out>0.02) %.4f'%(d3>0.02).mean(), flush=True)
    return d3
d3=run(iterations=10,featuredeg=30,checksurfdist=False)
bad=np.where(d3[:len(V)]>0.02)[0]
# which faces
vf=np.zeros(len(V),int); 
fv={}
for fi,t in zip(F,T):
    for v in t: fv.setdefault(v,set()).add(fi)
from collections import Counter
c=Counter()
for v in bad: 
    for f in fv[v]: c[(f,tab[f]['type'], round(tab[f].get('radius',0),2))]+=1
print(c.most_common(8))
run(iterations=10,featuredeg=30,checksurfdist=False,adaptive=True)
run(iterations=10,featuredeg=60,checksurfdist=False)
run(iterations=20,featuredeg=30,checksurfdist=False, smoothflag=False)
