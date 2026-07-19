import pandas as pd

# Leer el Excel
df = pd.read_excel(r"C:\Users\israe\Downloads\contactosm.xlsx")

print(df.columns)   # Ver los nombres de las columnas

with open("contactosm.vcf", "w", encoding="utf-8") as f:

    for _, fila in df.iterrows():

        nombre = str(fila["Nombre"]).strip()

        telefono = str(fila["Telefono"]).strip()
        telefono = telefono.replace(" ", "")
        telefono = telefono.replace("-", "")

        # Si Excel lo leyó como número decimal (593999999999.0)
        if telefono.endswith(".0"):
            telefono = telefono[:-2]

        f.write("BEGIN:VCARD\n")
        f.write("VERSION:3.0\n")
        f.write(f"N:;{nombre};;;\n")   # <-- todo el texto va en un solo campo, sin separar
        f.write(f"FN:{nombre}\n")
        f.write(f"TEL;TYPE=CELL:{telefono}\n")
        f.write("END:VCARD\n")

print("✅ Archivo contactos.vcf generado correctamente.")