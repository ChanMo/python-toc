import subprocess as sp
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def update_toc(f):
    """使用Liboffice更新TOC后保存
    @Parameter f: FilePath
    @Return: Boolean
    """
    res = sp.run(['soffice', '--headless', 'macro:///Mdian.AutoToc.UpdateIndexes({})'.format(f)], check=True)
    return res


def insert_toc(d, levels="1-3"):
    """
    Insert "Table of Contents" to Document

    Parameters:
    ------------

    d: Document Object
       文档对象

    levels: string
            default "1-3"
    根据 addheading 更新目录
    """
    sdt = OxmlElement('w:sdt')
    sdtpr = OxmlElement('w:sdtPr')
    docpartobj = OxmlElement('w:docPartObj')
    docpartgallery = OxmlElement('w:docPartGallery')
    docpartgallery.set(qn('w:val'), 'Table of Contents')
    docpartunique = OxmlElement('w:docPartUnique')
    docpartunique.set(qn('w:val'), 'true')
    docpartobj.append(docpartgallery)
    docpartobj.append(docpartunique)
    sdtpr.append(docpartobj)
    sdt.append(sdtpr)

    sdtcontent = OxmlElement('w:sdtContent')

    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = 'Contents'
    r.append(t)
    p.append(r)
    sdtcontent.append(p)

    fldChar = OxmlElement('w:fldChar')  # creates a new element
    fldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')  # sets attribute on element
    instrText.text = f'TOC \\o "{levels}" \\h \\z \\u'   # change 1-3 depending on heading levels you need

    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    # fldChar3 = OxmlElement('w:t')
    # fldChar3.text = "Right-click to update field."
    fldChar3 = OxmlElement('w:updateFields')
    fldChar3.set(qn('w:val'), 'true')
    fldChar2.append(fldChar3)

    fldChar4 = OxmlElement('w:fldChar')
    fldChar4.set(qn('w:fldCharType'), 'end')

    p2 = OxmlElement('w:p')
    r2 = OxmlElement('w:r')
    r2.append(fldChar)
    r2.append(instrText)
    r2.append(fldChar2)
    r2.append(fldChar4)
    p2.append(r2)

    sdtcontent.append(p2)
    sdt.append(sdtcontent)
    d._element.body.insert_element_before(sdt, *('w:sectPr',))

    return d

