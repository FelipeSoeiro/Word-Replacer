import os
import docx

def substituir_texto(dir_path, search_text, replace_text):
    # itera sobre todos os arquivos .docx na pasta especificada pelo usuário
    for filename in os.listdir(dir_path):
        if filename.endswith('.docx'):
            doc_path = os.path.join(dir_path, filename)
            doc = docx.Document(doc_path)

            # itera sobre todos os parágrafos do arquivo
            for para in doc.paragraphs:
                if search_text in para.text:
                    # substitui o texto encontrado pelo texto inserido pelo usuário
                    para.text = para.text.replace(search_text, replace_text)

            # salva o arquivo com as modificações
            doc.save(doc_path)

    print("Substituição concluída com sucesso!")

if __name__ == "__main__":
    # pede ao usuário o caminho da pasta a ser verificada
    dir_path = input("Insira o caminho da pasta a ser verificada: ")

    # pede ao usuário o texto a ser buscado e substituído
    search_text = input("Insira o texto a ser buscado e substituído: ")
    replace_text = input("Insira o novo texto: ")

    substituir_texto(dir_path, search_text, replace_text)
    