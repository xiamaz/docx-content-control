from _typeshed import ProfileFunction
from py_docx_cc import get_content_controls, map_content_controls


with open("./lb_complex.docx", "rb") as f:
    data = f.read()

controls = get_content_controls(data)

for key, meta in controls.items():
    print(key)
    print(meta.children_tags)
    print(meta.types)

with open("./lb_complex.docx", "rb") as f:
    data = f.read()


mapped = map_content_controls(data, {}, {"Hauptbefund": [{"Gen": "ABC1"}, {"Gen": "ABC2"}]})

with open("./lb_complex_mapped.docx", "wb") as f:
    f.write(mapped)
