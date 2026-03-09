from main import BOMGenerator

components = [
    {"reference": "C133", "value": "0.1uF", "package": "0603"},
    {"reference": "C138", "value": "20pF", "package": "0603"},
    {"reference": "C137", "value": "20pF", "package": "0603"},
    {"reference": "R44", "value": "12K", "package": "0603"},
    {"reference": "X1", "value": "30MHz", "package": ""},
    {"reference": "U12", "value": "CH395L", "package": ""},
]

bom_gen = BOMGenerator()
bom_df = bom_gen.generate_bom(components, merge=True)
bom_gen.display_bom(bom_df)
