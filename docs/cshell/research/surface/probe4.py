import sys,time,numpy as np, igl
d=np.load('docs/cshell/research/surface/cant_tess.npz');V,T,F=d['V'],d['T'],d['F']
import pymeshlab as pm
from scipy.sparse import coo_matrix
def dev(V2,T2):
    E=np.concatenate([V2[T2[:,[0,1]]].mean(1),V2[T2[:,[1,2]]].mean(1),V2[T2[:,[2,0]]].mean(1),V2[T2].mean(1)])
    d2=igl.point_mesh_squared_distance(E,V,T)[0]**.5
    nf=len(T2); fdev=d2.reshape(4,nf).max(0)
    Ci=np.concatenate([V,V[T].mean(1)]); d3,I,_=igl.point_mesh_squared_distance(Ci,V2,T2)
    d3=d3**.5
    np.maximum.at(fdev,I,d3)
    return fdev,d3.max()
ms=pm.MeshSet(); ms.add_mesh(pm.Mesh(V,T))
t0=time.time()
ms.meshing_isotropic_explicit_remeshing(targetlen=pm.PureValue(0.75),iterations=10,featuredeg=30,checksurfdist=False)
tol=0.02; L=0.75
for k in range(6):
    m=ms.current_mesh();V2=m.vertex_matrix();T2=m.face_matrix()
    fdev,hd=dev(V2,T2)
    bad=fdev>tol
    print(k,'nv',len(V2),'hd %.4f'%max(hd,fdev.max()),'bad faces',bad.sum(),'t %.0f'%(time.time()-t0),flush=True)
    if not bad.any(): break
    # dilate by 2 rings
    A=coo_matrix((np.ones(3*len(T2)*2),(np.repeat(np.arange(len(T2)),6),np.concatenate([T2[:,0],T2[:,1],T2[:,2],T2[:,0],T2[:,1],T2[:,2]]))),shape=(len(T2),len(V2))).tocsr()
    sel=bad.astype(float)
    for _ in range(2):
        vsel=(A.T@sel>0).astype(float); sel=((A@vsel)>0).astype(float)
    # local length: halve successively
    ms2=pm.MeshSet(); ms2.add_mesh(pm.Mesh(V2,T2,f_scalar_array=sel))
    ms2.compute_selection_by_condition_per_face(condselect="fq>0.5")
    L=L*0.6
    ms2.meshing_isotropic_explicit_remeshing(targetlen=pm.PureValue(L),iterations=5,featuredeg=30,selectedonly=True,checksurfdist=False)
    ms=ms2
