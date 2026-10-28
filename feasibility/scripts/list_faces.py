import sys, numpy as np
from OCP.BRepAdaptor import BRepAdaptor_Surface
from OCP.GProp import GProp_GProps
from OCP.BRepGProp import BRepGProp
from occ_util import DATA_DIR, OUT_DIR, load_step, faces, surf_type
s = load_step(f'{DATA_DIR}/{sys.argv[1]}.step')
for i, f in enumerate(faces(s)):
    t = surf_type(f)
    if t == 'BSplineSurface': continue
    g = GProp_GProps(); BRepGProp.SurfaceProperties_s(f, g); c = g.CentreOfMass(); A = g.Mass()
    a = BRepAdaptor_Surface(f); extra = ''
    if t == 'Cylinder':
        cy = a.Cylinder(); d = cy.Axis().Direction(); extra = f'r={cy.Radius():.2f} axis=({d.X():.2f},{d.Y():.2f},{d.Z():.2f})'
    if t == 'Plane':
        n = a.Plane().Axis().Direction(); extra = f'n=({n.X():.2f},{n.Y():.2f},{n.Z():.2f}) rev={f.Orientation()}'
    if A > 20 or t == 'Cylinder':
        print(i, t, f'A={A:.0f}', f'c=({c.X():.1f},{c.Y():.1f},{c.Z():.1f})', extra)
