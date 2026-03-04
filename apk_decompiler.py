# apk_decompiler.py - 支持AAB转APK版本（带日志回调）
import os
import subprocess
import shutil
import glob
import time
import zipfile
from colorama import Fore, Style, init
import re

init(autoreset=True)


class APKDecompiler:
    """APK/AAB资源文件反编译工具 - 只提取和反编译resources.arsc"""

    def __init__(self, log_callback=None):
        """
        初始化反编译器
        :param log_callback: 日志回调函数 callback(message, level='info')
                            level可以是: 'info', 'success', 'warning', 'error', 'primary'
        """
        self.log_callback = log_callback  # 必须在最前面，因为后续方法会使用
        self.apk_path = None
        self.output_dir = "decompiled_res"
        self.temp_apk_dir = "temp_minimal_apk"
        self.apktool_path = self._find_apktool()
        self.bundletool_path = self._find_bundletool()
        self.converted_apk = None  # 用于存储从AAB转换来的APK路径

    def _log(self, message, level='info'):
        """统一的日志输出方法"""
        # 移除colorama颜色代码
        clean_message = re.sub(r'\x1b\[[0-9;]*m', '', message)

        if self.log_callback:
            self.log_callback(clean_message, level)
        else:
            # 默认打印到控制台
            print(message)

    def _find_apktool(self):
        """查找apktool工具"""
        # 检查系统是否安装了apktool
        try:
            result = subprocess.run(['apktool', '--version'],
                                    capture_output=True,
                                    text=True,
                                    timeout=5)
            if result.returncode == 0:
                self._log("✓ 检测到已安装的apktool", 'success')
                return 'apktool'
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

        # 检查当前目录是否有apktool.jar
        if os.path.exists('apktool.jar'):
            self._log("✓ 检测到当前目录的apktool.jar", 'success')
            return 'apktool.jar'

        self._log("⚠ 未检测到apktool工具", 'warning')
        return None

    def _find_bundletool(self):
        """查找bundletool工具（用于AAB转APK）"""
        # 检查当前目录是否有bundletool.jar
        if os.path.exists('bundletool.jar'):
            self._log("✓ 检测到当前目录的bundletool.jar", 'success')
            return 'bundletool.jar'

        # 检查常见的bundletool文件名
        for name in ['bundletool-all.jar', 'bundletool-all-*.jar']:
            files = glob.glob(name)
            if files:
                self._log(f"✓ 检测到bundletool: {files[0]}", 'success')
                return files[0]

        return None

    def select_apk(self):
        """选择要反编译的APK文件"""
        # 如果已经设置了apk_path（比如从GUI传入），直接使用
        if self.apk_path and os.path.exists(self.apk_path):
            self._log(f"✓ 使用指定的APK/AAB文件: {self.apk_path}", 'success')
            return True

        # 搜索APK和AAB文件
        apk_files = glob.glob("**/*.apk", recursive=True) + glob.glob("**/*.aab", recursive=True)

        if not apk_files:
            self._log("✗ 未找到APK/AAB文件！", 'error')
            return False

        if len(apk_files) == 1:
            self.apk_path = apk_files[0]
            self._log(f"✓ 自动选择唯一的APK/AAB文件: {self.apk_path}", 'success')
            return True

        # 多个APK文件，让用户选择
        self._log("\n发现多个APK/AAB文件：", 'primary')
        apk_dict = {idx: apk for idx, apk in enumerate(apk_files, 1)}
        for idx, apk in apk_dict.items():
            size = os.path.getsize(apk) / (1024 * 1024)  # 转换为MB
            file_type = "AAB" if apk.endswith('.aab') else "APK"
            self._log(f"  {idx}. [{file_type}] {apk} ({size:.2f} MB)", 'info')

        while True:
            try:
                choice = input(f"\n请选择APK/AAB文件序号 (1-{len(apk_files)}): ")
                self.apk_path = apk_dict[int(choice)]
                self._log(f"✓ 已选择: {self.apk_path}", 'success')
                return True
            except (ValueError, KeyError):
                self._log(f"✗ 无效输入，请输入 1-{len(apk_files)} 之间的数字", 'error')
                continue

    def check_apktool(self):
        """检查并指导安装apktool"""
        if self.apktool_path:
            return True

        self._log("=" * 80, 'error')
        self._log("未找到apktool工具！", 'error')
        self._log("\n请按以下步骤安装apktool：", 'warning')
        self._log("\n方法1: 使用包管理器安装", 'primary')
        self._log("  Windows (Chocolatey): choco install apktool", 'info')
        self._log("  macOS (Homebrew):     brew install apktool", 'info')
        self._log("  Linux (apt):          sudo apt-get install apktool", 'info')
        self._log("\n方法2: 手动下载", 'primary')
        self._log("  1. 访问: https://ibotpeaches.github.io/Apktool/", 'info')
        self._log("  2. 下载 apktool.jar", 'info')
        self._log("  3. 将 apktool.jar 放在当前目录", 'info')
        self._log("=" * 80, 'error')

        return False

    def check_bundletool(self):
        """检查并指导安装bundletool"""
        if self.bundletool_path:
            return True

        self._log("=" * 80, 'error')
        self._log("未找到bundletool工具！", 'error')
        self._log("\nAAB文件需要先使用bundletool转换为APK", 'warning')
        self._log("\n下载bundletool：", 'primary')
        self._log("  1. 访问: https://github.com/google/bundletool/releases", 'info')
        self._log("  2. 下载最新的 bundletool-all-x.x.x.jar", 'info')
        self._log("  3. 重命名为 bundletool.jar 并放在当前目录", 'info')
        self._log("=" * 80, 'error')

        return False

    def is_aab_file(self):
        """判断是否为AAB文件"""
        return self.apk_path and self.apk_path.lower().endswith('.aab')

    def convert_aab_to_apk(self):
        """将AAB文件转换为通用APK"""
        if not self.is_aab_file():
            return self.apk_path

        self._log("=" * 80, 'primary')
        self._log("检测到AAB文件，开始转换为APK...", 'primary')
        self._log("=" * 80, 'primary')

        # 检查bundletool
        if not self.check_bundletool():
            return None

        try:
            # 创建临时目录
            if os.path.exists(self.temp_apk_dir):
                shutil.rmtree(self.temp_apk_dir)
            os.makedirs(self.temp_apk_dir)

            # 输出路径
            apks_output = os.path.join(self.temp_apk_dir, "output.apks")

            self._log("\n步骤1: 使用bundletool生成通用APK...", 'primary')

            # 构建bundletool命令（生成通用APK）
            cmd = [
                'java', '-jar', self.bundletool_path,
                'build-apks',
                f'--bundle={self.apk_path}',
                f'--output={apks_output}',
                '--mode=universal'
            ]

            self._log(f"执行命令: {' '.join(cmd)}", 'info')

            # 执行转换
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1
            )

            # 实时输出进度
            for line in process.stdout:
                self._log(line.rstrip(), 'info')

            process.wait()

            if process.returncode != 0:
                self._log(f"\n✗ AAB转APK失败！返回码: {process.returncode}", 'error')
                return None

            self._log("  ✓ APKS文件已生成", 'success')

            # 从.apks中提取universal.apk
            self._log("\n步骤2: 从APKS中提取通用APK...", 'primary')

            universal_apk_path = os.path.join(self.temp_apk_dir, "universal.apk")

            with zipfile.ZipFile(apks_output, 'r') as zip_ref:
                # .apks文件中包含universal.apk
                if 'universal.apk' in zip_ref.namelist():
                    zip_ref.extract('universal.apk', self.temp_apk_dir)
                    self._log("  ✓ 已提取universal.apk", 'success')
                else:
                    self._log("✗ APKS中未找到universal.apk", 'error')
                    return None

            # 检查文件大小
            apk_size = os.path.getsize(universal_apk_path) / (1024 * 1024)
            self._log(f"  ✓ 转换后的APK大小: {apk_size:.2f} MB", 'success')

            # 保存转换后的APK路径
            self.converted_apk = universal_apk_path

            self._log("\n✓ AAB已成功转换为APK！", 'success')
            return universal_apk_path

        except FileNotFoundError as e:
            self._log(f"\n✗ 错误: {e}", 'error')
            self._log("请确保Java已安装（运行bundletool需要Java）", 'error')
            return None
        except Exception as e:
            self._log(f"\n✗ AAB转APK过程出错: {e}", 'error')
            import traceback
            self._log(traceback.format_exc(), 'error')
            return None

    def create_minimal_apk(self, source_apk):
        """从原APK中删除无关文件，创建只包含资源的精简APK"""
        try:
            self._log("\n步骤: 解压APK并筛选资源文件...", 'primary')

            # APK的标准结构
            keep_patterns = [
                'resources.arsc',
                'AndroidManifest.xml',
                'res/',
            ]
            exclude_patterns = [
                '.dex',
                '.so',
                'lib/',
                'assets/',
                'META-INF/',
                'kotlin/',
                'okhttp3/',
            ]

            kept_files = []
            skipped_size = 0

            # 创建临时提取目录
            extract_dir = os.path.join(self.temp_apk_dir, "extracted")
            if os.path.exists(extract_dir):
                shutil.rmtree(extract_dir)
            os.makedirs(extract_dir)

            with zipfile.ZipFile(source_apk, 'r') as zip_ref:
                all_files = zip_ref.namelist()
                total_files = len(all_files)

                self._log(f"  正在分析 {total_files} 个文件...", 'info')

                for idx, file_name in enumerate(all_files, 1):
                    # 每100个文件输出一次进度
                    if idx % 100 == 0 or idx == total_files:
                        self._log(f"  进度: {idx}/{total_files} ({idx * 100 // total_files}%)", 'info')

                    # 检查是否需要保留
                    should_keep = False
                    for pattern in keep_patterns:
                        if file_name == pattern or file_name.startswith(pattern):
                            should_keep = True
                            break

                    # 检查是否需要排除
                    if should_keep:
                        for pattern in exclude_patterns:
                            if pattern in file_name or file_name.endswith(pattern):
                                should_keep = False
                                break

                    if should_keep:
                        try:
                            zip_ref.extract(file_name, extract_dir)
                            kept_files.append(file_name)
                        except:
                            pass
                    else:
                        # 统计跳过的文件大小
                        try:
                            info = zip_ref.getinfo(file_name)
                            skipped_size += info.file_size
                        except:
                            pass

            self._log(f"  ✓ 保留文件: {len(kept_files)} 个", 'success')
            self._log(f"  • 跳过大小: {skipped_size / (1024 * 1024):.2f} MB", 'warning')

            # 创建精简APK
            self._log("\n步骤: 创建精简APK（仅包含资源）...", 'primary')
            minimal_apk_path = os.path.join(self.temp_apk_dir, "minimal.apk")

            with zipfile.ZipFile(minimal_apk_path, 'w', zipfile.ZIP_DEFLATED) as minimal_zip:
                for idx, file_name in enumerate(kept_files, 1):
                    if idx % 50 == 0 or idx == len(kept_files):
                        self._log(f"  打包进度: {idx}/{len(kept_files)} ({idx * 100 // len(kept_files)}%)", 'info')

                    file_path = os.path.join(extract_dir, file_name)
                    if os.path.exists(file_path):
                        minimal_zip.write(file_path, file_name)

            minimal_size = os.path.getsize(minimal_apk_path) / 1024
            original_size = os.path.getsize(source_apk) / (1024 * 1024)
            reduction = (1 - minimal_size / (original_size * 1024)) * 100

            self._log(f"  ✓ 精简APK已创建: {minimal_size:.2f} KB", 'success')
            self._log(f"  • 原APK大小: {original_size:.2f} MB", 'info')
            self._log(f"  • 体积减小: {reduction:.1f}%", 'success')

            return minimal_apk_path

        except Exception as e:
            self._log(f"✗ 创建精简APK失败: {e}", 'error')
            import traceback
            self._log(traceback.format_exc(), 'error')
            return None

    def decompile(self):
        """反编译APK文件（删除无关文件后反编译）"""
        if not self.check_apktool():
            return False

        if not self.select_apk():
            return False

        # 如果是AAB文件，先转换为APK
        working_apk = self.apk_path
        if self.is_aab_file():
            working_apk = self.convert_aab_to_apk()
            if not working_apk:
                self._log("✗ AAB转APK失败，无法继续", 'error')
                return False
        else:
            self._log("=" * 80, 'primary')
            self._log("开始创建精简APK并反编译...", 'primary')
            self._log("=" * 80, 'primary')

        # 创建临时目录（如果还没创建）
        if not os.path.exists(self.temp_apk_dir):
            os.makedirs(self.temp_apk_dir)

        # 创建精简APK（删除无关文件）
        minimal_apk = self.create_minimal_apk(working_apk)
        if not minimal_apk:
            self._log("✗ 无法创建精简APK，反编译失败", 'error')
            return False

        # 清理旧的输出目录
        if os.path.exists(self.output_dir):
            self._log("\n⚠ 删除旧的反编译目录...", 'warning')
            shutil.rmtree(self.output_dir)

        try:
            self._log("\n步骤: 反编译精简APK...", 'primary')

            # 构建apktool命令 - 反编译精简APK
            # 不使用-r -s参数，让apktool正常反编译资源
            if self.apktool_path == 'apktool.jar':
                cmd = ['java', '-jar', 'apktool.jar', 'd', minimal_apk,
                       '-o', self.output_dir, '-f']
            else:
                cmd = ['apktool', 'd', minimal_apk,
                       '-o', self.output_dir, '-f']

            self._log(f"执行命令: {' '.join(cmd)}", 'info')
            self._log("参数说明: 反编译精简APK（已删除dex、lib等无关文件）\n", 'info')

            # 执行反编译
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1
            )

            # 实时输出进度
            for line in process.stdout:
                self._log(line.rstrip(), 'info')

            process.wait()

            if process.returncode != 0:
                self._log(f"\n✗ 反编译失败！返回码: {process.returncode}", 'error')
                return False

            # 检查res目录
            res_dir = os.path.join(self.output_dir, 'res')
            if not os.path.exists(res_dir):
                self._log("\n✗ 反编译完成，但未找到res目录！", 'error')
                return False

            # 清理临时APK目录
            self._log("\n步骤: 清理临时文件...", 'primary')
            if os.path.exists(self.temp_apk_dir):
                shutil.rmtree(self.temp_apk_dir)
                self._log("  ✓ 已清理临时文件", 'success')

            # 统计结果
            strings_files = glob.glob(os.path.join(res_dir, '**/strings.xml'), recursive=True)

            self._log("\n" + "=" * 80, 'success')
            self._log("✓ 反编译成功！", 'success')
            self._log("=" * 80, 'success')
            self._log(f"\n资源文件位置: {res_dir}", 'primary')
            self._log(f"找到 {len(strings_files)} 个strings.xml文件\n", 'primary')

            # 按语言分组显示
            language_files = {}
            for sf in strings_files:
                rel_path = os.path.relpath(sf, res_dir)
                # 提取语言代码
                parts = rel_path.split(os.sep)
                if len(parts) > 0:
                    lang = parts[0]  # values, values-zh-rCN等
                    file_size = os.path.getsize(sf) / 1024
                    language_files[lang] = file_size

            # 排序并显示
            for lang in sorted(language_files.keys()):
                file_size = language_files[lang]
                self._log(f"  - {lang}/strings.xml ({file_size:.2f} KB)", 'info')

            # 显示最终大小
            total_size = sum(os.path.getsize(os.path.join(root, f))
                             for root, _, files in os.walk(self.output_dir)
                             for f in files) / (1024 * 1024)  # MB
            self._log(f"\n💾 反编译文件总大小: {total_size:.2f} MB", 'success')
            if self.is_aab_file():
                self._log("⚡ AAB已转换为APK并成功反编译！", 'success')
            else:
                self._log("⚡ 精简方案：删除无关文件，快速反编译！", 'success')

            return True

        except FileNotFoundError as e:
            self._log(f"\n✗ 错误: {e}", 'error')
            self._log("请确保Java已安装（运行apktool.jar需要Java）", 'error')
            return False
        except Exception as e:
            self._log(f"\n✗ 反编译过程出错: {e}", 'error')
            import traceback
            self._log(traceback.format_exc(), 'error')
            return False
        finally:
            # 确保清理临时文件
            if os.path.exists(self.temp_apk_dir):
                try:
                    shutil.rmtree(self.temp_apk_dir)
                except:
                    pass

    def get_res_directory(self):
        """获取反编译后的res目录路径"""
        res_dir = os.path.join(self.output_dir, 'res')
        if os.path.exists(res_dir):
            return res_dir
        return None

    def cleanup(self):
        """清理反编译产生的文件"""
        cleaned = False

        if os.path.exists(self.output_dir):
            try:
                shutil.rmtree(self.output_dir)
                self._log("✓ 已清理反编译文件", 'success')
                cleaned = True
            except Exception as e:
                self._log(f"⚠ 清理反编译文件失败: {e}", 'warning')

        if os.path.exists(self.temp_apk_dir):
            try:
                shutil.rmtree(self.temp_apk_dir)
                if not cleaned:  # 避免重复打印
                    self._log("✓ 已清理临时文件", 'success')
            except Exception as e:
                self._log(f"⚠ 清理临时APK文件失败: {e}", 'warning')


def main():
    """独立运行时的测试函数"""
    print(f"{Fore.CYAN}{'=' * 80}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}APK/AAB资源文件反编译工具{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{'=' * 80}{Style.RESET_ALL}\n")

    decompiler = APKDecompiler()

    if decompiler.decompile():
        res_dir = decompiler.get_res_directory()
        print(f"\n{Fore.GREEN}可以开始使用res目录进行后续操作: {res_dir}{Style.RESET_ALL}")

        # 询问是否清理
        choice = input(f"\n{Fore.YELLOW}是否清理反编译文件？(y/n): {Style.RESET_ALL}").lower()
        if choice == 'y':
            decompiler.cleanup()
    else:
        print(f"\n{Fore.RED}反编译失败，请检查错误信息{Style.RESET_ALL}")

    input("\nPress Enter to exit...")


if __name__ == '__main__':
    main()