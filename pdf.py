# -*- coding: utf-8 -*-
# -------------------------------
# @软件：PyCharm
# @PyCharm：自行填入你的版本号
# @Python：自行填入你的版本号
# @项目：PythonProject2
# -------------------------------
# @文件：pdf.py
# @时间：2025/8/12 23:07
# -------------------------------
# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
完整的PDF转DOCX转换器
使用pdf2docx库，支持格式保持和批量转换
作者：Assistant
版本：1.0
"""

import os
import sys
import time
import subprocess
from pathlib import Path
from typing import List, Tuple, Optional
import logging


class PDFtoDocxConverter:
    """PDF转DOCX转换器类"""

    def __init__(self, log_level=logging.INFO):
        """初始化转换器"""
        self.setup_logging(log_level)
        self.logger = logging.getLogger(__name__)
        self._check_dependencies()

    def setup_logging(self, level):
        """设置日志"""
        logging.basicConfig(
            level=level,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.StreamHandler(),
                logging.FileHandler('pdf_converter.log', encoding='utf-8')
            ]
        )

    def _check_dependencies(self):
        """检查并安装必要的依赖"""
        try:
            from pdf2docx import Converter
            self.logger.info("✅ pdf2docx库已安装")
        except ImportError:
            self.logger.warning("⚠️ pdf2docx库未安装，正在自动安装...")
            self._install_pdf2docx()

    def _install_pdf2docx(self):
        """自动安装pdf2docx库"""
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pdf2docx"])
            self.logger.info("✅ pdf2docx库安装成功")
            # 重新导入
            from pdf2docx import Converter
        except subprocess.CalledProcessError as e:
            self.logger.error(f"❌ 安装pdf2docx失败: {e}")
            self.logger.error("请手动运行: pip install pdf2docx")
            sys.exit(1)

    def convert_single_file(self, pdf_path: str, docx_path: str,
                            start_page: int = 0, end_page: int = None) -> bool:
        """
        转换单个PDF文件到DOCX

        Args:
            pdf_path: PDF文件路径
            docx_path: 输出DOCX文件路径
            start_page: 起始页码（从0开始）
            end_page: 结束页码（None表示到最后一页）

        Returns:
            bool: 转换是否成功
        """
        try:
            from pdf2docx import Converter

            # 验证输入文件
            if not os.path.exists(pdf_path):
                self.logger.error(f"❌ PDF文件不存在: {pdf_path}")
                return False

            # 创建输出目录
            output_dir = os.path.dirname(docx_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
                self.logger.info(f"📁 创建输出目录: {output_dir}")

            self.logger.info(f"🔄 开始转换: {os.path.basename(pdf_path)}")
            start_time = time.time()

            # 执行转换
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=start_page, end=end_page)
            cv.close()

            # 计算转换时间
            elapsed_time = time.time() - start_time

            # 检查输出文件
            if os.path.exists(docx_path):
                file_size = os.path.getsize(docx_path) / (1024 * 1024)  # MB
                self.logger.info(f"✅ 转换成功: {os.path.basename(docx_path)}")
                self.logger.info(f"📊 文件大小: {file_size:.2f} MB")
                self.logger.info(f"⏱️ 转换耗时: {elapsed_time:.2f} 秒")
                return True
            else:
                self.logger.error(f"❌ 转换失败，输出文件未创建")
                return False

        except Exception as e:
            self.logger.error(f"❌ 转换过程中发生错误: {str(e)}")
            return False

    def batch_convert(self, input_dir: str, output_dir: str = None,
                      file_pattern: str = "*.pdf") -> Tuple[int, int]:
        """
        批量转换PDF文件

        Args:
            input_dir: 输入目录
            output_dir: 输出目录（默认为输入目录_converted）
            file_pattern: 文件匹配模式

        Returns:
            tuple: (成功数量, 总数量)
        """
        input_path = Path(input_dir)

        if not input_path.exists():
            self.logger.error(f"❌ 输入目录不存在: {input_dir}")
            return (0, 0)

        # 设置输出目录
        if output_dir is None:
            output_dir = f"{input_dir}_converted"

        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        # 查找所有PDF文件
        pdf_files = list(input_path.glob(file_pattern))
        if not pdf_files:
            self.logger.warning(f"⚠️ 在 {input_dir} 中没有找到匹配的PDF文件")
            return (0, 0)

        self.logger.info(f"📁 找到 {len(pdf_files)} 个PDF文件")
        self.logger.info(f"📤 输出目录: {output_dir}")

        success_count = 0
        start_time = time.time()

        for i, pdf_file in enumerate(pdf_files, 1):
            docx_file = output_path / f"{pdf_file.stem}.docx"

            self.logger.info(f"📄 处理文件 [{i}/{len(pdf_files)}]: {pdf_file.name}")

            # 检查输出文件是否已存在
            if docx_file.exists():
                response = input(f"文件 {docx_file.name} 已存在，是否覆盖？(y/n/a=全部/s=跳过全部): ").lower()
                if response in ['n', 'no']:
                    self.logger.info(f"⏭️ 跳过: {pdf_file.name}")
                    continue
                elif response in ['s', 'skip']:
                    self.logger.info("⏭️ 跳过所有现有文件")
                    break
                elif response in ['a', 'all']:
                    pass  # 继续覆盖所有

            if self.convert_single_file(str(pdf_file), str(docx_file)):
                success_count += 1

            # 显示进度
            progress = (i / len(pdf_files)) * 100
            self.logger.info(f"📈 进度: {progress:.1f}% ({success_count}/{i} 成功)")

        # 批量转换完成统计
        total_time = time.time() - start_time
        self.logger.info(f"\n🎉 批量转换完成!")
        self.logger.info(f"📊 成功: {success_count}/{len(pdf_files)} 个文件")
        self.logger.info(f"⏱️ 总耗时: {total_time:.2f} 秒")

        return (success_count, len(pdf_files))

    def get_pdf_info(self, pdf_path: str) -> Optional[dict]:
        """
        获取PDF文件信息

        Args:
            pdf_path: PDF文件路径

        Returns:
            dict: PDF信息字典
        """
        try:
            from pdf2docx import Converter
            import fitz  # PyMuPDF，pdf2docx的依赖

            doc = fitz.open(pdf_path)
            info = {
                'pages': doc.page_count,
                'title': doc.metadata.get('title', ''),
                'author': doc.metadata.get('author', ''),
                'subject': doc.metadata.get('subject', ''),
                'creator': doc.metadata.get('creator', ''),
                'file_size': f"{os.path.getsize(pdf_path) / (1024 * 1024):.2f} MB"
            }
            doc.close()

            return info
        except Exception as e:
            self.logger.error(f"❌ 获取PDF信息失败: {e}")
            return None

    def interactive_mode(self):
        """交互式模式"""
        print("\n" + "=" * 50)
        print("🔄 PDF转DOCX转换器 - 交互模式")
        print("=" * 50)

        while True:
            print("\n请选择操作:")
            print("1. 转换单个PDF文件")
            print("2. 批量转换PDF文件")
            print("3. 查看PDF文件信息")
            print("4. 退出程序")

            choice = input("\n请输入选择 (1-4): ").strip()

            if choice == '1':
                self._interactive_single_convert()
            elif choice == '2':
                self._interactive_batch_convert()
            elif choice == '3':
                self._interactive_pdf_info()
            elif choice == '4':
                print("👋 谢谢使用!")
                break
            else:
                print("❌ 无效选择，请重试")

    def _interactive_single_convert(self):
        """交互式单文件转换"""
        pdf_path = input("📄 请输入PDF文件路径: ").strip().strip('"')

        if not os.path.exists(pdf_path):
            print(f"❌ 文件不存在: {pdf_path}")
            return

        # 默认输出路径
        default_docx = os.path.splitext(pdf_path)[0] + ".docx"
        docx_path = input(f"💾 输出DOCX文件路径 (默认: {default_docx}): ").strip().strip('"')

        if not docx_path:
            docx_path = default_docx

        # 页码范围（可选）
        page_range = input("📖 页码范围 (格式: 开始页-结束页，如 1-5，留空转换全部): ").strip()
        start_page, end_page = 0, None

        if page_range:
            try:
                if '-' in page_range:
                    start, end = page_range.split('-')
                    start_page = int(start) - 1  # 转为0基索引
                    end_page = int(end)
                else:
                    start_page = int(page_range) - 1
                    end_page = start_page + 1
            except ValueError:
                print("⚠️ 页码格式无效，将转换全部页面")

        # 执行转换
        success = self.convert_single_file(pdf_path, docx_path, start_page, end_page)
        if success:
            print(f"🎉 转换完成: {docx_path}")
        else:
            print("❌ 转换失败")

    def _interactive_batch_convert(self):
        """交互式批量转换"""
        input_dir = input("📁 请输入PDF文件所在目录: ").strip().strip('"')

        if not os.path.exists(input_dir):
            print(f"❌ 目录不存在: {input_dir}")
            return

        output_dir = input(f"💾 输出目录 (默认: {input_dir}_converted): ").strip().strip('"')
        if not output_dir:
            output_dir = f"{input_dir}_converted"

        success, total = self.batch_convert(input_dir, output_dir)
        print(f"\n🎉 批量转换完成: {success}/{total} 个文件成功")

    def _interactive_pdf_info(self):
        """交互式PDF信息查看"""
        pdf_path = input("📄 请输入PDF文件路径: ").strip().strip('"')

        info = self.get_pdf_info(pdf_path)
        if info:
            print(f"\n📋 PDF文件信息:")
            print(f"📖 页数: {info['pages']}")
            print(f"📝 标题: {info['title'] or '未设置'}")
            print(f"👤 作者: {info['author'] or '未设置'}")
            print(f"📄 主题: {info['subject'] or '未设置'}")
            print(f"🔧 创建工具: {info['creator'] or '未设置'}")
            print(f"💾 文件大小: {info['file_size']}")


def main():
    """主函数"""
    # 创建转换器实例
    converter = PDFtoDocxConverter()

    # 检查命令行参数
    if len(sys.argv) == 1:
        # 没有参数，启动交互模式
        converter.interactive_mode()
    else:
        # 命令行模式
        if len(sys.argv) < 2:
            print("使用方法:")
            print("  python pdf2docx_complete.py                    # 交互模式")
            print("  python pdf2docx_complete.py input.pdf          # 转换单个文件")
            print("  python pdf2docx_complete.py input.pdf output.docx  # 指定输出文件")
            print("  python pdf2docx_complete.py /path/to/pdfs batch    # 批量转换")
            return

        if sys.argv[-1] == 'batch' or os.path.isdir(sys.argv[1]):
            # 批量转换模式
            input_dir = sys.argv[1]
            output_dir = sys.argv[2] if len(sys.argv) > 2 and sys.argv[2] != 'batch' else None
            converter.batch_convert(input_dir, output_dir)
        else:
            # 单文件转换模式
            pdf_path = sys.argv[1]
            docx_path = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(pdf_path)[0] + ".docx"

            success = converter.convert_single_file(pdf_path, docx_path)
            if success:
                print(f"🎉 转换完成: {docx_path}")
            else:
                print("❌ 转换失败")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⏹️ 用户中断操作")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ 程序发生错误: {e}")
        sys.exit(1)