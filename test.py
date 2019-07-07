import odfmetachanger as OMC


et = OMC.read_odf_meta_data("test.odt")

fm = OMC.load_yaml_frontmatter("test.md")
for key, value in fm.items():
    if key == "title":
        OMC.alter_odf_meta_title(et, value)
    else:
        OMC.alter_odf_meta_user(et, key, value)

OMC.create_new_odf_file("test.odt", et)




#et.write("test.xml", encoding="UTF-8", xml_declaration=True)
