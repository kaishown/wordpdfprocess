#enconding=utf-8
#author yanggx
#time 2021/4/29

from docx import Document
def textRepalce(document_file,save_path):
    """
    文本框中的文字ori_text替换为replace_text
    :param document_file: str
    :param save_path: str
    :return: None
    """
    ori_text="gguixiu"
    replace_text = "杨桂秀"
    file = Document(document_file)
    children = file.element.body.iter()
    child_iters = []
    tags = []
    for child in children:
        # 通过类型判断目录
        if child.tag.endswith(('AlternateContent', 'textbox')):
            for indx, ci in enumerate(child.iter()):
                tags.append(ci.tag)
                if ci.tag.endswith(('main}r', 'main}pPr')):
                    child_iters.append(ci)
    text = ['']
    for ci in child_iters:
        if ci.tag.endswith('main}pPr'):
            text.append('')
        else:
            text[-1] += ci.text
        if ci.text == ori_text:  # 只能替换ci，不等于文本框的所有内容
            ci.text =replace_text
    file.save(save_path)


if __name__ == '__main__':
    document_file = 'exu.docx'
    save_path = "save_path.docx"
    textRepalce(document_file, save_path)