# %%
from pptx import Presentation

pptx_path = 'templates/template.pptx'

prs = Presentation(pptx_path)

# %%

# prs.slides[3].shapes[4].table.cell(0,0).text = 'this is my test'


# %%
slide = prs.slides[3]

# %%

for shape in slide.shapes:
    if shape.shape_type == 'table':
        print(shape)
# %%
shapes = []

for shape in slide.shapes:
    if shape.has_table:
        shapes.append(shape)

shapes
shape_to_change = shapes[1]
shape_to_change.table.cell(0,0).text = "did this work" # this does work
# %%
# %%
prs.slides[3].shapes[0].table.cell(0,0).text = 'this is shape 0, (0,0)'
prs.slides[3].shapes[1].table.cell(0,0).text = 'this is shape 1 (0,0)'


prs.save('test.pptx')
