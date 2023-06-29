# -*- coding: utf-8 -*-
"""
Created on Tue Jun 27 11:02:16 2023

@author: 86136
"""

"""
实现功能：文档解析
输入：PDF
输出：文档目录、文档内图片（PNG格式）、文档表格（采用pandas的dataFrame数据类型表示）、可返回的文档版面数据
"""

import argparse
import pdfplumber
import pdfminer.high_level
import fitz
import pandas as pd
import os
import PyPDF2
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument


def pdf_extract_tables(pdf_path):
    """
    功能：从pdf中提取表格
    输入：pdf文件路径pdf_path（"example.pdf"）
    输出：对于每一个表输出一个pandas的dataFrame类型，采用print方式
    返回值：表格df
    """
    
    # open pdf with pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                
                # for every table extract data to dataFrame
                df = pd.DataFrame(table[1:], columns=table[0])
                print(df)
        pdf.close()
    return df


def pdf_extract_images(pdf_path, pic_path):

    """
    功能：从pdf中提取图片（从pdf文件中提取所有图片，并将它们保存为png文件）
    输入：pdf文件路径pdf_path（"example.pdf"），图片保存地址result_path
    输出：对于每一个图片以png格式直接输出到result_path
    返回值：无
    """

    pdf = os.path.join(pdf_path)
    print(pdf)

    # 打开pdf，打印PDF的相关信息
    doc = fitz.open(pdf)
    # 图片计数
    imgcount = 0
    lenXREF = doc.xref_length()    #获取pdf文件对象总数

    # 打印PDF的信息
    print("文件名:{}, 页数: {}, 对象: {}".format(pdf, len(doc), lenXREF - 1))

    if not os.path.exists(pic_path):
        os.makedirs(pic_path)
    # 遍历doc，获取每一页
    for page in doc:
        try:
            imgcount +=1
            tupleImage = page.get_images()
            for xref in list(tupleImage):
                xref = list(xref)[0]
                img = doc.extract_image(xref)   #获取文件扩展名，图片内容 等信息
                imageFilename = ("%s-%s." % (imgcount, xref) + img["ext"])
                imageFilename = imageFilename  #合成最终 的图像的文件名
                imageFilename = os.path.join(pic_path, imageFilename)   #合成最终图像完整路径名
                print(imageFilename)
                imgout = open(imageFilename, 'wb')   #byte方式新建图片
                imgout.write(img["image"])   #当前提取的图片写入磁盘
                imgout.close
        except:
            continue


def pdf_get_toc(pdf_path):
    """
    功能：从pdf中提取目录并输出
    输入：pdf文件路径pdf_path（"example.pdf"）
    输出：一个list对象，包括每个章节的标题
    返回值：章节的标题list
    """

    # open pdf file and new a PDFParser object
    fp = open(pdf_path, 'rb')
    parser = PDFParser(fp)
    document = PDFDocument(parser)
    try:
        # get outlines
        outlines = document.get_outlines()
    except:
        raise Exception("PDF文档没有目录,请检查PDF生成时的设置")
    result = []
    
    # append title in result list
    for (level, title, dest, a, se) in outlines:
        result.append(title)
    print(result)
    return result
    

def get_pdf_text(pdf_path):
    """
    功能：从pdf中提取正文并输出
    输入：pdf文件路径pdf_path（"example.pdf"）
    输出：输出正文内容
    返回值：正文
    """
    
    # extract the plaintext
    text = pdfminer.high_level.extract_text(pdf_path)
    print(text)
    return text


if __name__ == "__main__":
    # add a parser
    parser = argparse.ArgumentParser(description='PDF Extraction')
    
    # add cmd args
    parser.add_argument('-o','--option', type=str, default='toc', help='The function will be served, optional choice will be text(Extract text), image(Extract images), table(Extract tables) and toc(Extract tocs)')
    parser.add_argument('-p','--path', type=str, default='demo.pdf', help='Path to the pdf file')
    parser.add_argument('-r','--result_path', type=str, default='./pic', help='Path to store the result images')
    
    # parse the args
    args = parser.parse_args()
    
    # parse the options
    if args.option == 'text':
        get_pdf_text(args.path)
    elif args.option == 'image':
        pdf_extract_images(args.path, args.result_path)
    elif args.option == 'table':
        pdf_extract_tables(args.path)
    elif args.option == 'toc':
        pdf_get_toc(args.path)
    else:
        print("Invalid options.")
