import base64
import os
import glob

output_file = 'tera_assets.py'
png_files = glob.glob('*.png')

with open(output_file, 'w') as out:
    out.write('"""\nAuto-generated file containing base64 encoded PNG assets for the TERA Report Generator.\n"""\n\n')
    for png in png_files:
        with open(png, 'rb') as f:
            b64_data = base64.b64encode(f.read()).decode('utf-8')
            var_name = os.path.splitext(png)[0].upper().replace('-', '_').replace(' ', '_').replace(']', '')
            out.write(f'{var_name} = "{b64_data}"\n\n')
print(f"Generated {output_file} from {len(png_files)} PNG files.")
