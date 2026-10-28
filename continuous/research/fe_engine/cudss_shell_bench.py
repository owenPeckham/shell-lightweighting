# cuDSS benchmark: Q1 hex elasticity (E=1,nu=.3,h=1) on a thin-walled hollow box (wall t elements), z=0 face clamped.
# usage: python cudss_shell_bench.py L t [hybrid]
import sys, time, numpy as np, cupy as cp, cupyx.scipy.sparse as csp
import nvmath.sparse.advanced as nsa
L=int(sys.argv[1]); t=int(sys.argv[2]); hyb=len(sys.argv)>3
def ke():
    E,nu=1.0,0.3; lam=E*nu/((1+nu)*(1-2*nu)); mu=E/(2*(1+nu))
    D=np.zeros((6,6)); D[:3,:3]=lam; D[range(3),range(3)]+=2*mu; D[range(3,6),range(3,6)]=mu
    s=np.array([[-1,-1,-1],[1,-1,-1],[1,1,-1],[-1,1,-1],[-1,-1,1],[1,-1,1],[1,1,1],[-1,1,1]],float)
    g=1/np.sqrt(3); K=np.zeros((24,24))
    for gx in(-g,g):
     for gy in(-g,g):
      for gz in(-g,g):
        dN=np.array([[s[i,0]*(1+s[i,1]*gy)*(1+s[i,2]*gz),s[i,1]*(1+s[i,0]*gx)*(1+s[i,2]*gz),s[i,2]*(1+s[i,0]*gx)*(1+s[i,1]*gy)] for i in range(8)])/8*2  # d/dx with h=1 -> J=1/2
        B=np.zeros((6,24))
        for i in range(8):
            B[0,3*i]=dN[i,0];B[1,3*i+1]=dN[i,1];B[2,3*i+2]=dN[i,2]
            B[3,3*i]=dN[i,1];B[3,3*i+1]=dN[i,0];B[4,3*i+1]=dN[i,2];B[4,3*i+2]=dN[i,1];B[5,3*i]=dN[i,2];B[5,3*i+2]=dN[i,0]
        K+=B.T@D@B*0.125
    return K
KE=cp.asarray(ke())
T0=time.time()
i,j,k=cp.meshgrid(cp.arange(L),cp.arange(L),cp.arange(L),indexing='ij')
m=((i<t)|(i>=L-t)|(j<t)|(j>=L-t)|(k<t)|(k>=L-t)); ex,ey,ez=i[m],j[m],k[m]; ne=int(ex.size); del i,j,k,m
nid=lambda x,y,z:(x*(L+1)+y)*(L+1)+z
en=cp.stack([nid(ex+a,ey+b,ez+c) for a,b,c in [(0,0,0),(1,0,0),(1,1,0),(0,1,0),(0,0,1),(1,0,1),(1,1,1),(0,1,1)]],1)
used=cp.unique(en.ravel()); nn=int(used.size); mp=cp.full((L+1)**3,-1,cp.int64); mp[used]=cp.arange(nn)
en=mp[en]; X=(used//((L+1)**2)); Z=used%(L+1)
free=cp.repeat(Z>0,3); fid=cp.cumsum(free)-1; dof_map=cp.where(free,fid,-1); nf=int(free.sum())
edof=(en[:,:,None]*3+cp.arange(3)).reshape(ne,24); edof=dof_map[edof]
A=None; ch=150000
for s0 in range(0,ne,ch):
    ed=edof[s0:s0+ch]; n=ed.shape[0]
    r=cp.repeat(ed[:,:,None],24,2).ravel(); c=cp.repeat(ed[:,None,:],24,1).ravel(); v=cp.tile(KE.ravel(),n)
    ok=(r>=0)&(c>=0); Ac=csp.coo_matrix((v[ok],(r[ok],c[ok])),shape=(nf,nf)).tocsr()
    A=Ac if A is None else A+Ac
A=A.tocsr(); A.sort_indices(); cp.cuda.Device().synchronize()
print(f"L={L} t={t} elems={ne} nodes={nn} free_dof={nf} nnz={A.nnz} assemble={time.time()-T0:.1f}s",flush=True)
del en,edof,KE
rng=cp.random.default_rng(0); b=rng.standard_normal((nf,1))
def mem(): f,tot=cp.cuda.runtime.memGetInfo(); return (tot-f)/2**30
opts=nsa.DirectSolverOptions(sparse_system_type=nsa.DirectSolverMatrixType.SPD,multithreading_lib='/opt/sa/lib/python3.12/site-packages/nvidia/cu12/lib/libcudss_mtlayer_gomp.so.0')
ex_=nsa.ExecutionCUDA(hybrid_memory_mode_options=nsa.HybridMemoryModeOptions(hybrid_memory_mode=True)) if hyb else None
cp.get_default_memory_pool().free_all_blocks(); m0=mem(); print(f"GPU mem before solver {m0:.1f} GiB (matrix {(A.data.nbytes+A.indices.nbytes+A.indptr.nbytes)/2**30:.1f} GiB) hybrid={hyb}",flush=True)
s=nsa.DirectSolver(A,b,options=opts,execution=ex_)
t0=time.time(); s.plan_config.host_nthreads=72; s.plan(); cp.cuda.Device().synchronize(); tp=time.time()-t0; print(f"plan(reorder+symbolic) {tp:.1f}s mem {mem():.1f}",flush=True)
t0=time.time(); s.factorize(); cp.cuda.Device().synchronize(); tf=time.time()-t0; print(f"factorize {tf:.2f}s mem {mem():.1f} GiB",flush=True)
t0=time.time(); x=s.solve(); cp.cuda.Device().synchronize(); ts=time.time()-t0; 
t0=time.time(); x=s.solve(); cp.cuda.Device().synchronize(); ts=time.time()-t0
res=float(cp.linalg.norm(A@x[:,0]-b[:,0])/cp.linalg.norm(b)); print(f"solve(1rhs) {ts:.3f}s relres {res:.1e}",flush=True)
# refactor with same pattern (scaled values)
A.data*=1.5; s.reset_operands(a=A)
t0=time.time(); s.factorize(); cp.cuda.Device().synchronize(); print(f"refactorize(same pattern) {time.time()-t0:.2f}s",flush=True)
A.data/=1.5; s.reset_operands(a=A)
s.factorize()
# multi-RHS block solve (Lanczos/LOBPCG-block use)
for nr in (6,12):
    try:
        B=rng.standard_normal((nf,nr)); s.reset_operands(b=B); s.solve()
        cp.cuda.Device().synchronize(); t0=time.time(); X=s.solve(); cp.cuda.Device().synchronize(); print(f"solve({nr} rhs) {time.time()-t0:.3f}s",flush=True)
    except Exception as e: print("multi-rhs fail",nr,str(e)[:100]); break
s.free()
