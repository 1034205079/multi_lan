# apk_decompiler.py - APK反编译工具（带日志回调 + Java支持：优先系统，备用项目）
import os
import subprocess
import shutil
import glob
import time
import zipfile
from colorama import Fore, Style, init
import re
import platform

init(autoreset=True)


class APKDecompiler:
    """APK资源文件反编译工具 - 只提取和反编译resources.arsc"""

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
        self.java_path = self._find_java()
        self.apktool_path = self._find_apktool()

    def _log(self, message, level='info'):
        """统一的日志输出方法"""
        # 移除colorama颜色代码
        clean_message = re.sub(r'\x1b\[[0-9;]*m', '', message)

        if self.log_callback:
            self.log_callback(clean_message, level)
        else:
            # 默认打印到控制台
            print(message)

    def _find_java(self):
        """查找Java环境（优先使用系统Java，系统没有再用项目内的）"""
        # 1. 检查系统Java环境（优先）
        try:
            result = subprocess.run(['java', '-version'],
                                    capture_output=True,
                                    text=True,
                                    timeout=5)
            if result.returncode == 0 or 'java version' in result.stderr.lower():
                self._log("✓ 使用系统Java环境", 'success')
                return 'java'
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

        # 2. 检查项目内的Java（备用）
        script_dir = os.path.dirname(os.path.abspath(__file__))

        # 根据操作系统确定Java可执行文件名
        if platform.system() == 'Windows':
            java_exe = 'java.exe'
        else:
            java_exe = 'java'

        # 项目内Java路径
        local_java_paths = [
            os.path.join(script_dir, 'java', 'bin', java_exe),
            os.path.join(script_dir, 'jre', 'bin', java_exe),
            os.path.join(script_dir, 'jdk', 'bin', java_exe),
        ]

        for java_path in local_java_paths:
            if os.path.exists(java_path):
                try:
                    result = subprocess.run([java_path, '-version'],
                                            capture_output=True,
                                            text=True,
                                            timeout=5)
                    if result.returncode == 0 or 'java version' in result.stderr.lower():
                        self._log(f"✓ 使用项目内的Java（系统未安装）: {java_path}", 'success')
                        return java_path
                except Exception:
                    pass

        self._log("⚠ 未检测到Java环境（系统和项目内都没有）", 'warning')
        return None

    def _find_apktool(self):
        """查找apktool工具（优先使用项目内的apktool）"""
        # 1. 检查项目内的apktool.jar
        script_dir = os.path.dirname(os.path.abspath(__file__))

        local_apktool_paths = [
            os.path.join(script_dir, 'apktool', 'apktool.jar'),
            os.path.join(script_dir, 'apktool.jar'),
        ]

        for apktool_path in local_apktool_paths:
            if os.path.exists(apktool_path):
                self._log(f"✓ 检测到项目内的apktool: {apktool_path}", 'success')
                return apktool_path

        # 2. 检查系统是否安装了apktool
        try:
            result = subprocess.run(['apktool', '--version'],
                                    capture_output=True,
                                    text=True,
                                    timeout=5)
            if result.returncode == 0:
                self._log("✓ 检测到系统安装的apktool", 'success')
                return 'apktool'
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

        self._log("⚠ 未检测到apktool工具", 'warning')
        return None

    def select_apk(self):
        """选择要反编译的APK文件"""
        # 如果已经设置了apk_path（比如从GUI传入），直接使用
        if self.apk_path and os.path.exists(self.apk_path):
            self._log(f"✓ 使用指定的APK文件: {self.apk_path}", 'success')
            return True

        # 搜索APK文件
        apk_files = glob.glob("**/*.apk", recursive=True)

        if not apk_files:
            self._log("✗ 未找到APK文件！", 'error')
            return False

        if len(apk_files) == 1:
            self.apk_path = apk_files[0]
            self._log(f"✓ 自动选择唯一的APK文件: {self.apk_path}", 'success')
            return True

        # 多个APK文件，让用户选择
        self._log("\n发现多个APK文件：", 'primary')
        apk_dict = {idx: apk for idx, apk in enumerate(apk_files, 1)}
        for idx, apk in apk_dict.items():
            size = os.path.getsize(apk) / (1024 * 1024)  # 转换为MB
            self._log(f"  {idx}. {apk} ({size:.2f} MB)", 'info')

        while True:
            try:
                choice = input(f"\n请选择APK文件序号 (1-{len(apk_files)}): ")
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
        self._log("  3. 将 apktool.jar 放在 apktool 文件夹中", 'info')
        self._log("=" * 80, 'error')

        return False

    def check_java(self):
        """检查Java环境"""
        if self.java_path:
            return True

        self._log("=" * 80, 'error')
        self._log("未找到Java环境！", 'error')
        self._log("\n请确保：", 'warning')
        self._log("  1. 系统已安装Java（JRE或JDK）（推荐）", 'info')
        self._log("  2. 或者项目内包含 java 文件夹（备用）", 'info')
        self._log("=" * 80, 'error')

        return False

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
        if not self.check_java():
            return False

        if not self.check_apktool():
            return False

        if not self.select_apk():
            return False

        # 使用APK文件
        working_apk = self.apk_path
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
            # 判断apktool是jar文件还是系统命令
            if self.apktool_path.endswith('.jar'):
                # 使用jar文件
                cmd = [self.java_path, '-jar', self.apktool_path, 'd', minimal_apk,
                       '-o', self.output_dir, '-f']
            else:
                # 使用系统命令
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
            self._log("⚡ 精简方案：删除无关文件，快速反编译！", 'success')

            return True

        except FileNotFoundError as e:
            self._log(f"\n✗ 错误: {e}", 'error')
            self._log("请确保Java已正确配置", 'error')
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
    print(f"{Fore.CYAN}APK资源文件反编译工具{Style.RESET_ALL}")
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