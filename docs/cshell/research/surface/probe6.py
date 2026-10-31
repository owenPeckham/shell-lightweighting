import numpy as np, igl
d=np.load('docs/cshell/research/surface/cant_tess.npz');V,T,F=d['V'],d['T'],d['F']
pd1,pd2,pv1,pv2=igl.principal_curvatures(V,T,5) if hasattr(igl,'principal_curvatures') else (None,)*4
