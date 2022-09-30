import pytest

from main import BrowserAction
from time import sleep
import json
from tools import *

driver = BrowserAction()
pro_path = os.path.dirname(os.path.dirname(__file__))
path = get_conf('FCS', '测试结果文件夹')
fcs_result_path = pro_path + "/" + path


class TestCase:

    def test_convert0_001(self):
        """参数convertTimeOut为60s，转换超时时间60s"""
        data1 = eval(get_cell(fcs_result_path, 2, 4, "Sheet0-ss"))
        data2 = eval(get_cell(fcs_result_path, 2, 6, 'Sheet0-ss'))
        set_time = data1["convertTimeOut"]  # 设置转换的超时时间，秒
        convert_time = int(data2["data"]["convertTime"]) / 1000  # 实际转换实际，毫秒
        result = data2["message"]  # 转换结果
        if set_time >= convert_time:  # 设置的超时时间大于等于实际转换时间，应该转换成功
            assert result == "操作成功"
        else:  # 设置的超时时间小于实际转换时间，应该转换超时
            assert result == "转换超时，请调整转换超时时间"

    def test_convert0_002(self):
        """参数allowFileSize，允许转换的文件大小，单位M"""
        file_path = get_cell(fcs_result_path, 3, 2, 'Sheet0-ss')  # 获取文件路径
        file_size = get_file_size(file_path)  # 获取实际文档大小
        data = eval(get_cell(fcs_result_path, 3, 4, 'Sheet0-ss'))  # 获取允许转换的文件大小
        allow_size = data["allowFileSize"]
        result = get_cell(fcs_result_path, 3, 7, 'Sheet0-ss')  # 获取文件实际转换截图
        if file_size <= allow_size:  # 如果实际文件size小于等于允许转换的文件大小，则判断转换结果为”通过“
            assert result == "通过"
        else:  # 如果实际文件size大于允许转换的文件大小，则判断转换结果为”失败“
            assert result == "失败"

    def test_convert0_003(self):
        if get_cell(fcs_result_path, 4, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数htmlName，转换后的文档标题名称"""
        url = get_cell(fcs_result_path, 4, 8, 'Sheet0-ss')
        data = eval(get_cell(fcs_result_path, 4, 4, 'Sheet0-ss'))
        html_name = data['htmlName']  # 转换后的预期标题名称
        driver.open_bro(url)
        name = driver.get_text("//*[@class='title']")  # 转换后的实际标题名称
        assert name == html_name  # 验证转换后的实际标题名称等于预期的

    def test_convert0_004(self):
        if get_cell(fcs_result_path, 5, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数htmlTitle，转换后的html标签名称"""
        url = get_cell(fcs_result_path, 5, 8, 'Sheet0-ss')
        data = eval(get_cell(fcs_result_path, 5, 4, 'Sheet0-ss'))
        html_title = data['htmlTitle']  # 转换后html预期标签名称
        driver.open_bro(url)
        assert driver.get_title() == html_title  # 验证转换后html实际标签等于预期的

    def test_convert0_005(self):
        if get_cell(fcs_result_path, 6, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数isCopy，0不防复制，可以打开右键菜单"""
        url = get_cell(fcs_result_path, 6, 8, 'Sheet0-ss')
        driver.open_bro(url)
        driver.open_context_menu(driver.find_ele("//div[@id='content']"))
        sleep(2)
        """截图验证"""
        driver.screenshot_save(6, 'Sheet0-ss')

    def test_convert0_006(self):
        if get_cell(fcs_result_path, 7, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数isCopy，1防复制，不可以打开右键菜单"""
        url = get_cell(fcs_result_path, 7, 8, 'Sheet0-ss')
        driver.open_bro(url)
        driver.open_context_menu(driver.find_ele("//div[@id='content']"))
        sleep(2)
        """截图验证"""
        driver.screenshot_save(7, 'Sheet0-ss')

    def test_convert0_007(self):
        if get_cell(fcs_result_path, 8, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数page，转换页码，page为几，就转第几页"""
        url = get_cell(fcs_result_path, 8, 8, 'Sheet0-ss')
        driver.open_bro(url)
        total_page = int(driver.get_text("//*[@class='totalPage']"))  # 实际共转换的页数
        assert 1 == total_page  # 验证只转1页
        """截图验证"""
        driver.screenshot_save(8, 'Sheet0-ss')

    def test_convert0_008(self):
        if get_cell(fcs_result_path, 9, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数pageStart，pageEnd，从第n页转到第m页"""
        url = get_cell(fcs_result_path, 9, 8, 'Sheet0-ss')
        data = eval(get_cell(fcs_result_path, 9, 4, 'Sheet0-ss'))
        page_start = int(data['pageStart'])
        page_end = int(data['pageEnd'])
        page = page_end - page_start + 1  # 预期转换总页数
        driver.open_bro(url)
        total_page = int(driver.get_text("//*[@class='totalPage']"))  # 实际转换的页数
        assert total_page == page  # 实际转换的页数等于预期的
        """截图验证"""
        driver.screenshot_save(9, 'Sheet0-ss')

    def test_convert0_009(self):
        if get_cell(fcs_result_path, 10, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数page,pageStart，pageEnd组合"""
        url = get_cell(fcs_result_path, 10, 8, 'Sheet0-ss')
        driver.open_bro(url)
        total_page = int(driver.get_text("//*[@class='totalPage']"))  # 实际共转换的页数
        assert 1 == total_page  # 验证只转1页
        """截图验证"""
        driver.screenshot_save(10, "Sheet0-ss")

    def test_convert0_010(self):
        if get_cell(fcs_result_path, 11, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数password不加密文档密码正确123456"""
        data = eval(get_cell(fcs_result_path, 11, 6, 'Sheet0-ss'))
        result = data["message"]  # 转换结果
        assert result == "操作成功"  # 参数输入正确密码，验证转换成功
        url = get_cell(fcs_result_path, 11, 8, 'Sheet0-ss')
        driver.open_bro(url)
        """打开转换后的链接，进行截图验证"""
        driver.screenshot_save(11, 'Sheet0-ss')

    def test_convert0_011(self):
        if get_cell(fcs_result_path, 12, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数password加密文档密码正确123456"""
        data = eval(get_cell(fcs_result_path, 12, 6, 'Sheet0-ss'))
        result = data["message"]  # 转换结果
        assert result == "操作成功"  # 参数输入正确密码，验证转换成功
        url = get_cell(fcs_result_path, 12, 8, 'Sheet0-ss')
        driver.open_bro(url)
        """打开转换后的链接，进行截图验证"""
        driver.screenshot_save(12, 'Sheet0-ss')

    def test_convert0_012(self):
        """参数password加密文档密码错误123"""
        data = eval(get_cell(fcs_result_path, 13, 6, 'Sheet0-ss'))
        result = data["message"]  # 转换结果
        assert result == "转换的文档为加密文档或密码有误，请重新添加password参数进行转换"  # 参数输入错误密码，验证转换失败

    def test_convert0_013(self):
        if get_cell(fcs_result_path, 14, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数download，参数为0，不显示源文档下载按钮"""
        url = get_cell(fcs_result_path, 14, 8, 'Sheet0-ss')
        driver.open_bro(url)
        assert not driver.ele_exist("//img[@title='下载']")  # 验证参数为0时，没有“下载”按钮存在

    def test_convert0_014(self):
        if get_cell(fcs_result_path, 15, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数download，参数为1，显示源文档下载按钮,点击下载后，验证项目download文件夹下有该文件"""
        url = get_cell(fcs_result_path, 15, 8, 'Sheet0-ss')
        driver.open_bro(url)
        del_download_file()  # 删除下载目录,重新新建
        file_name = driver.get_text("//span[@class='title']")  # 获取文件名
        assert driver.ele_exist("//img[@title='下载']")  # 验证参数为1时，有“下载”按钮存在
        driver.click_ele("//img[@title='下载']")
        sleep(2)
        assert (download_file_exist(file_name))  # 验证项目download文件夹下有该文件

    def test_convert0_015(self):
        if get_cell(fcs_result_path, 16, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数is_signature为0，不进入签批模式"""
        url = get_cell(fcs_result_path, 16, 8, 'Sheet0-ss')
        driver.open_bro(url)
        assert not driver.ele_exist("//img[@title='签批']")  # 验证没有签批按钮存在

    def test_convert0_016(self):
        if get_cell(fcs_result_path, 17, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数is_signature为1，进入签批模式"""
        url = get_cell(fcs_result_path, 17, 8, 'Sheet0-ss')
        driver.open_bro(url)
        assert driver.ele_exist("//img[@title='签批']")  # 验证有签批按钮存在
        driver.click_ele("//img[@title='签批']")  # 点击签批按钮
        sleep(1)
        """截图验证"""
        driver.screenshot_save(17, 'Sheet0-ss')

    def test_convert0_017(self):
        if get_cell(fcs_result_path, 18, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数isHeaderBar为0，不显示头部导航栏"""
        url = get_cell(fcs_result_path, 18, 8, 'Sheet0-ss')
        driver.open_bro(url)
        assert not driver.ele_exist("//div[@id='header']")  # 验证头部导航栏这个元素不存在

    def test_convert0_018(self):
        if get_cell(fcs_result_path, 19, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """参数isHeaderBar为1，显示头部导航栏"""
        url = get_cell(fcs_result_path, 19, 8, 'Sheet0-ss')
        driver.open_bro(url)
        assert driver.ele_exist("//div[@id='header']")  # 验证头部导航栏这个元素存在

    def test_convert0_019(self):
        if get_cell(fcs_result_path, 20, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """num参数为0，预览次数不限"""
        url = get_cell(fcs_result_path, 20, 8, 'Sheet0-ss')
        for i in range(1, 5):
            """预览url，打开4次，验证能否打开成功"""
            driver.open_bro(url)
            assert driver.ele_exist("//div[@id='root']")  # 验证主体栏存在
            i += 1

    def test_convert0_020(self):
        if get_cell(fcs_result_path, 21, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """num参数为3，打开预览超过3次不支持预览"""
        url = get_cell(fcs_result_path, 21, 8, 'Sheet0-ss')
        for i in range(1, 5):
            """预览url打开小于等于3次，验证能否打开成功"""
            driver.open_bro(url)
            if i <= 3:
                assert driver.ele_exist("//div[@id='root']")  # 验证主体栏存在
            else:
                data = json.loads(driver.get_text("//pre"))
                assert data['message'] == '链接已过期'
            i += 1

    def test_convert0_021(self):
        if get_cell(fcs_result_path, 22, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
        """"time参数为预览过期时间，打开预览time秒后，刷新页面，显示链接过期"""
        url = get_cell(fcs_result_path, 22, 8, 'Sheet0-ss')
        data = eval(get_cell(fcs_result_path, 22, 4, 'Sheet0-ss'))
        time_num = int(data["time"])  # 获取time数据,并转换成int类型
        driver.open_bro(url)
        assert driver.ele_exist("//div[@id='root']")  # 验证主体栏存在
        sleep(time_num)  # 预览打开time_num秒后自动刷新网页
        driver.open_bro(url)
        data = json.loads(driver.get_text("//pre"))
        assert data['message'] == '链接已过期'  # 预览时间到期后，显示“链接过期”

    def test_convert0_022(self):
        if get_cell(fcs_result_path, 23, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmContent显示水印内容"""
            url = get_cell(fcs_result_path, 23, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(23, 'Sheet0-ss')

    def test_convert0_023(self):
        if get_cell(fcs_result_path, 24, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmSize为10"""
            url = get_cell(fcs_result_path, 24, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(24, 'Sheet0-ss')

    def test_convert0_024(self):
        if get_cell(fcs_result_path, 25, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmSize为30"""
            url = get_cell(fcs_result_path, 25, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(25, 'Sheet0-ss')

    def test_convert0_025(self):
        if get_cell(fcs_result_path, 26, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmColor为#111111"""
            url = get_cell(fcs_result_path, 26, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(26, 'Sheet0-ss')

    def test_convert0_026(self):
        if get_cell(fcs_result_path, 27, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmFont为微软雅黑"""
            url = get_cell(fcs_result_path, 27, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(27, 'Sheet0-ss')

    def test_convert0_027(self):
        if get_cell(fcs_result_path, 28, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmMargin为30，40"""
            url = get_cell(fcs_result_path, 28, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(28, 'Sheet0-ss')

    def test_convert0_028(self):
        if get_cell(fcs_result_path, 29, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmMargin为100，200"""
            url = get_cell(fcs_result_path, 29, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(29, 'Sheet0-ss')

    def test_convert0_029(self):
        if get_cell(fcs_result_path, 30, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmTransparency为0.2"""
            url = get_cell(fcs_result_path, 30, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(30, 'Sheet0-ss')

    def test_convert0_030(self):
        if get_cell(fcs_result_path, 31, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmTransparency为0"""
            url = get_cell(fcs_result_path, 31, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(31, 'Sheet0-ss')

    def test_convert0_031(self):
        if get_cell(fcs_result_path, 32, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmTransparency为1"""
            url = get_cell(fcs_result_path, 32, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(32, 'Sheet0-ss')

    def test_convert0_032(self):
        if get_cell(fcs_result_path, 33, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmRotate为90"""
            url = get_cell(fcs_result_path, 33, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(33, 'Sheet0-ss')

    def test_convert0_033(self):
        if get_cell(fcs_result_path, 34, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmRotate为45"""
            url = get_cell(fcs_result_path, 34, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(34, 'Sheet0-ss')

    def test_convert0_034(self):
        if get_cell(fcs_result_path, 35, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数"wmPosition为1,1"""
            url = get_cell(fcs_result_path, 35, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(35, 'Sheet0-ss')

    def test_convert0_035(self):
        if get_cell(fcs_result_path, 36, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数"wmPosition为300,400"""
            url = get_cell(fcs_result_path, 36, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(36, 'Sheet0-ss')

    def test_convert0_036(self):
        if get_cell(fcs_result_path, 37, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmPicPath"""
            url = get_cell(fcs_result_path, 37, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(37, 'Sheet0-ss')

    def test_convert0_037(self):
        if get_cell(fcs_result_path, 38, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmPicSize为10"""
            url = get_cell(fcs_result_path, 38, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(38, 'Sheet0-ss')

    def test_convert0_038(self):
        if get_cell(fcs_result_path, 39, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmMargin为30，40"""
            url = get_cell(fcs_result_path, 39, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(39, 'Sheet0-ss')

    def test_convert0_039(self):
        if get_cell(fcs_result_path, 40, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmMargin为100，200"""
            url = get_cell(fcs_result_path, 40, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(40, 'Sheet0-ss')

    def test_convert0_040(self):
        if get_cell(fcs_result_path, 41, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmTransparency为0.2"""
            url = get_cell(fcs_result_path, 41, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(41, 'Sheet0-ss')

    def test_convert0_041(self):
        if get_cell(fcs_result_path, 42, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmTransparency为0"""
            url = get_cell(fcs_result_path, 42, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(42, 'Sheet0-ss')

    def test_convert0_042(self):
        if get_cell(fcs_result_path, 43, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmTransparency为1"""
            url = get_cell(fcs_result_path, 43, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(43, 'Sheet0-ss')

    def test_convert0_043(self):
        if get_cell(fcs_result_path, 44, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmRotate为45"""
            url = get_cell(fcs_result_path, 44, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(44, 'Sheet0-ss')

    def test_convert0_044(self):
        if get_cell(fcs_result_path, 45, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmRotate为90"""
            url = get_cell(fcs_result_path, 45, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(45, 'Sheet0-ss')

    def test_convert0_045(self):
        if get_cell(fcs_result_path, 46, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmPosition为1，1"""
            url = get_cell(fcs_result_path, 46, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(46, 'Sheet0-ss')

    def test_convert0_046(self):
        if get_cell(fcs_result_path, 47, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """参数wmPosition为300，400"""
            url = get_cell(fcs_result_path, 47, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(46, 'Sheet0-ss')

    def test_convert0_047(self):
        if get_cell(fcs_result_path, 48, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """水印内容组合"""
            url = get_cell(fcs_result_path, 48, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(48, 'Sheet0-ss')

    def test_convert0_048(self):
        if get_cell(fcs_result_path, 49, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """图片水印组合"""
            url = get_cell(fcs_result_path, 49, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(49, 'Sheet0-ss')

    def test_convert0_049(self):
        if get_cell(fcs_result_path, 50, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """无水印内容组合"""
            url = get_cell(fcs_result_path, 50, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(50, 'Sheet0-ss')

    def test_convert0_050(self):
        if get_cell(fcs_result_path, 51, 7, "Sheet0-ss") != '通过':
            pytest.skip("url转换失败,不执行该case")
            """无图片水印组合"""
            url = get_cell(fcs_result_path, 51, 8, 'Sheet0-ss')
            driver.open_bro(url)
            """截图验证"""
            driver.screenshot_save(51, 'Sheet0-ss')
