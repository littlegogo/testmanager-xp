# -*- coding:UTF-8 -*-
import io
import json
from Ui_cranetestdoc import Ui_CraneTestDocWnd
# from PyQt4 import QtGui, QtCore

from PyQt4.QtGui import QDialog, QFileDialog
from PyQt4.QtGui import QTreeWidgetItem, QTreeWidgetItemIterator, QMessageBox, QTableWidgetItem
from PyQt4.QtCore import Qt, QDateTime, QString
from testcase import TestCase
from docwriter import DocWriter

TEST_CAT_KEY = 'test_category'
TEST_PERSON_KEY = 'test_persons'
TEST_ENV_KEY = 'test_environment'
TEST_REQ_METHOD_KEY = 'qualified_method'


class CraneTestDocWnd(QDialog, Ui_CraneTestDocWnd):
    def __init__(self):
        self._case_id = 1
        self.test_cases = {}
        self.config = self.__load_config()
        self.init_ui()
        self.connect_slots()

    def init_ui(self):
        """
        init ui
        :return:
        """
        QDialog.__init__(self)
        Ui_CraneTestDocWnd.__init__(self)
        self.setupUi(self)
        self.testcase_tree.clear()
        # 初始化下拉框
        self.test_person_combox.addItems(self.config[TEST_PERSON_KEY])
        self.test_cat_combox.addItems(self.config[TEST_CAT_KEY])
        self.test_env_combox.addItems(self.config[TEST_ENV_KEY])
        self.qualified_method_combox.addItems(self.config[TEST_REQ_METHOD_KEY])
        # 设置表格宽度
        self.test_procedure_tabel.setColumnWidth(0, 50)
        self.test_procedure_tabel.setColumnWidth(1, 300)
        self.test_procedure_tabel.setColumnWidth(2, 300)
        self.test_procedure_tabel.setColumnWidth(3, 300)
        self.test_procedure_tabel.setColumnWidth(6, 0)
        # 设置时间选择框
        self.test_date_timepickedit.setDateTime(QDateTime.currentDateTime())
        # 初始化进度条
        self.process_progressbar.setValue(0)

        # 设置窗口大小
        self.resize(1700, 900)

    def connect_slots(self):
        self.testcase_tree.itemDoubleClicked.connect(self.on_testcase_tree_doubleclicked)
        self.testcase_tree.itemClicked.connect(self.on_testcase_tree_clicked)
        self.add_test_case_button.clicked.connect(self.on_add_testcase)
        self.save_test_case_button.clicked.connect(self.on_export_testcase)
        self.remove_test_case_button.clicked.connect(self.on_remove_testcase)
        self.generate_doc_button.clicked.connect(self.on_generate_doc)
        self.import_case_button.clicked.connect(self.on_import_case)
        self.selectall_checkbox.stateChanged.connect(self.on_select_all)
        # 添加/删除测试步骤按钮
        self.add_procedure_button.clicked.connect(self.on_add_test_procdure)
        self.remove_procdure_button.clicked.connect(self.on_remove_test_procdure)
        # 树形列表选择改变
        self.testcase_tree.currentItemChanged.connect(self.on_current_changed)

    def clear_case_content(self):
        """
        清除显示的内容
        :return:
        """
        self.test_caseidentify_edit.setText('')
        self.test_item_edit.setText('')
        self.test_cat_combox.setCurrentIndex(0)
        self.require_trace_edit.setText('')
        self.test_content_edit.setText('')
        self.sys_prepare_edit.setText('')
        self.precondation_edit.setText('')
        self.test_input_edit.setText('')
        for i in range(0, self.test_procedure_tabel.rowCount()):
            self.test_procedure_tabel.removeRow(0)
        self.estimate_rule_eidt.setText('')
        self.test_env_combox.setCurrentIndex(0)
        self.qualified_method_combox.setCurrentIndex(0)
        self.safe_secret_edit.setText('')
        self.test_person_combox.setCurrentIndex(0)
        self.test_person_join_edit.setText('')
        self.test_date_timepickedit.setDateTime(QDateTime.currentDateTime())
        self.test_data_edit.setText('')
        self.problem_sheet_edit.setText('')
        self.correct_sheet_edit.setText('')
        self.test_diff_edit.setText('')

    def update_display(self, test_case):
        self.test_caseidentify_edit.setText(unicode(test_case.case_mark))
        self.test_item_edit.setText(unicode(test_case.case_name))
        self.test_cat_combox.setCurrentIndex(test_case.case_cat)
        self.require_trace_edit.setText(unicode(test_case.case_req_track))
        self.test_content_edit.setText(unicode(test_case.case_content))
        self.sys_prepare_edit.setText(unicode(test_case.case_sys_prepare))
        self.precondation_edit.setText(unicode(test_case.case_constraint))
        self.test_input_edit.setText(unicode(test_case.case_input))
        for index, item in enumerate(test_case.case_exec_procedure):
            self.test_procedure_tabel.insertRow(index)
            self.test_procedure_tabel.setItem(index, 0, QTableWidgetItem(item[0]))
            self.test_procedure_tabel.setItem(index, 1, QTableWidgetItem(unicode(item[1])))
            self.test_procedure_tabel.setItem(index, 2, QTableWidgetItem(unicode(item[2])))
            self.test_procedure_tabel.setItem(index, 3, QTableWidgetItem(unicode(item[3])))
            self.test_procedure_tabel.setItem(index, 4, QTableWidgetItem(unicode(item[4])))
            self.test_procedure_tabel.setItem(index, 5, QTableWidgetItem(unicode(item[5])))
            self.test_procedure_tabel.setItem(index, 6, QTableWidgetItem(unicode(item[6])))
        self.estimate_rule_eidt.setText(unicode(test_case.case_qualified_rule))
        self.test_env_combox.setCurrentIndex(test_case.case_env)
        self.qualified_method_combox.setCurrentIndex(test_case.case_qualified_method)
        self.safe_secret_edit.setText(unicode(test_case.case_safe_secret))
        self.test_person_combox.setCurrentIndex(test_case.test_person)
        self.test_person_join_edit.setText(unicode(test_case.test_join_person))
        self.test_date_timepickedit.setDateTime(QDateTime.fromString(test_case.test_date, 'yyyy-MM-dd'))
        self.test_data_edit.setText(unicode(test_case.case_data))
        self.problem_sheet_edit.setText(unicode(test_case.case_problem_sheet))
        self.correct_sheet_edit.setText(unicode(test_case.case_correct_sheet))
        self.test_diff_edit.setText(unicode(test_case.case_diff))

    def on_current_changed(self, current, previous):
        if current and str(current.text(0)) in self.test_cases:
            self.clear_case_content()
            self.update_display(self.test_cases[str(current.text(0))])

    def on_testcase_tree_clicked(self, item):
        pass

    def on_select_all(self):
        if self.selectall_checkbox.isChecked():
            for i in range(0, self.testcase_tree.topLevelItemCount()):
                self.testcase_tree.topLevelItem(i).setCheckState(0, Qt.Checked)
        else:
            for i in range(0, self.testcase_tree.topLevelItemCount()):
                self.testcase_tree.topLevelItem(i).setCheckState(0, Qt.Unchecked)

    def on_testcase_tree_doubleclicked(self):
        pass

    def on_add_testcase(self):
        """
        添加用例
        :return:
        """
        test_case_identification = str(self.test_caseidentify_edit.text())
        test_case_name = str(self.test_item_edit.text())
        if test_case_identification and test_case_name:
            if test_case_identification not in self.test_cases:
                #  在树形列表中增加一项
                test_case = self.__generate_test_case(test_case_identification)
                item = QTreeWidgetItem()
                item.setText(0, test_case.case_mark)
                item.setText(1, test_case.case_name)
                item.setText(2, test_case.case_id)
                item.setCheckState(0, Qt.Checked)
                self.testcase_tree.addTopLevelItem(item)
            else:
                test_case = self.__generate_test_case(test_case_identification)
                # 更新树形列表中的用例名称
                for i in range(0, self.testcase_tree.topLevelItemCount()):
                    if self.testcase_tree.topLevelItem(i).text(0) == test_case.case_mark:
                        self.testcase_tree.topLevelItem(i).setText(1, test_case.case_name)
                QMessageBox.information(self, u'提示', u'用例[{0}({1})]已更新'.format(test_case_name, test_case.case_mark), QMessageBox.Yes, QMessageBox.Yes)
        else:
            QMessageBox.information(self, u'提示', u'用例标识和用例名称必填', QMessageBox.Yes, QMessageBox.Yes)
            return

    def on_export_testcase(self):
        """
        slot function for save test case button
        :return:
        """
        total_export_count=0
        for i in range(0, self.testcase_tree.topLevelItemCount()):
            if self.testcase_tree.topLevelItem(i).checkState(0) == Qt.Checked and str(self.testcase_tree.topLevelItem(i).text(0)) in self.test_cases:
                total_export_count += 1
        if not total_export_count:
            QMessageBox.information(self, u'提示', u'未选择任何需要导出的用例,在左侧列表中选择需要导出的用例',
                                    QMessageBox.Yes, QMessageBox.Yes)
            return

        export_dir = QFileDialog.getExistingDirectory(self, u'选择导出用例文件的目录', '.')
        if not export_dir:
            return
        self.process_progressbar.setMaximum(total_export_count)
        self.process_progressbar.setMinimum(0)
        export_count = 0
        try:
            for i in range(0, self.testcase_tree.topLevelItemCount()):
                if self.testcase_tree.topLevelItem(i).checkState(0) == Qt.Checked and str(self.testcase_tree.topLevelItem(i).text(0)) in self.test_cases:
                    test_case = self.test_cases[str(self.testcase_tree.topLevelItem(i).text(0))]

                    test_case.save_to_file('{0}/{1}.case'.format(str(export_dir), test_case.case_mark))
                    export_count = export_count + 1
                    self.process_progressbar.setValue(export_count)
            QMessageBox.information(self, u'提示', u'{0}/{1}个用例保存成功'.format(export_count, total_export_count), QMessageBox.Yes, QMessageBox.Yes)
        except Exception as e:
            QMessageBox.warning(self, u'错误', u'用例保存异常' + str(e), QMessageBox.Yes, QMessageBox.Yes)

    def on_remove_testcase(self):
        """
        slot function for remove test case button
        :return:
        """
        info = u'是否删除所选测试用例(已存储的用例文件请手动删除)!!!'
        if QMessageBox.Yes == QMessageBox.question(self, u'询问', info, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes):
            items = QTreeWidgetItemIterator(self.testcase_tree)
            while items.value():
                if items.value().checkState(0) == Qt.Checked:
                    self.test_cases.pop(str(items.value().text(0)))
                    self.testcase_tree.takeTopLevelItem(self.testcase_tree.indexOfTopLevelItem(items.value()))
                    continue
                items += 1

    def on_add_test_procdure(self):
        """
        添加测试步骤
        :return:
        """
        row_count = self.test_procedure_tabel.rowCount()
        self.test_procedure_tabel.insertRow(row_count)
        self.test_procedure_tabel.setItem(row_count, 0, QTableWidgetItem(str(row_count+1)))

    def on_remove_test_procdure(self):
        current_index = self.test_procedure_tabel.currentIndex()
        if current_index.row() >= 0:
            self.test_procedure_tabel.removeRow(current_index.row())
            for i in range(0, self.test_procedure_tabel.rowCount()):
                self.test_procedure_tabel.item(i, 0).setText(str(i+1))

    def on_import_case(self):
        query_result = QMessageBox.warning(self, u'警告', u'执行导入操作时，测试标识相同的用例将被覆盖，确认执行导入操作!?',
                            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if query_result == QMessageBox.No:
            return
        case_files = QFileDialog.getOpenFileNames(self, u'选择用例文件', '.', '')
        total_import_count = len(case_files)
        if not total_import_count:
            return
        self.process_progressbar.setMaximum(total_import_count)
        self.process_progressbar.setMinimum(0)
        fail = 0
        import_count = 0
        try:
            for file in case_files:
                test_case = TestCase()
                test_case.load_from_file(str(file))
                ret = self.testcase_tree.findItems(test_case.case_mark, Qt.MatchCaseSensitive | Qt.MatchExactly, 0)
                if ret:
                    ret[0].setText(1, unicode(test_case.case_name))
                else:
                    item = QTreeWidgetItem()
                    item.setText(0, unicode(test_case.case_mark))
                    item.setText(1, unicode(test_case.case_name))
                    item.setText(2, test_case.case_id)
                    item.setCheckState(0, Qt.Checked)
                    self.testcase_tree.addTopLevelItem(item)
                self.test_cases[test_case.case_mark] = test_case
                import_count += 1
                self.process_progressbar.setValue(import_count)
        except Exception as e:
            fail += 1
            QMessageBox.information(self, u'提示', u'导入用例文件异常{0}'.format(str(e)), QMessageBox.Yes, QMessageBox.Yes)
        QMessageBox.information(self, u'提示', u'导入成功{0}个, 失败{1}个'.format(import_count, fail), QMessageBox.Yes, QMessageBox.Yes)

    def __generate_test_case(self, case_identification=''):
        """
        生成或更新测试用例
        :param case_identification: 测试标识
        :return:
        """
        test_case = None
        try:
            test_case = self.test_cases[case_identification]
        except KeyError:
            self._case_id += 1
            test_case = TestCase(case_id=str(self._case_id), case_mark=case_identification)
            self.test_cases[case_identification] = test_case
        # 更新除用例id和用例标识之外的其它内容
        test_case.case_name = str(self.test_item_edit.text()).decode('utf-8')
        test_case.case_cat = self.test_cat_combox.currentIndex()
        test_case.case_req_track = str(self.require_trace_edit.text())
        test_case.case_content = str(self.test_content_edit.text())
        test_case.case_sys_prepare = str(self.sys_prepare_edit.text())
        test_case.case_constraint = str(self.precondation_edit.text())
        test_case.case_input = str(self.test_input_edit.text())
        test_case.case_exec_procedure[:] = []
        for i in range(0, self.test_procedure_tabel.rowCount()):
            test_case.case_exec_procedure.append([
                str(self.test_procedure_tabel.item(i, 0).text()) if self.test_procedure_tabel.item(i, 0) else str(i+1),
                str(self.test_procedure_tabel.item(i, 1).text()) if self.test_procedure_tabel.item(i, 1) else '无',
                str(self.test_procedure_tabel.item(i, 2).text()) if self.test_procedure_tabel.item(i, 2) else '无',
                str(self.test_procedure_tabel.item(i, 3).text()) if self.test_procedure_tabel.item(i, 3) else '无',
                str(self.test_procedure_tabel.item(i, 4).text()) if self.test_procedure_tabel.item(i, 4) else '无',
                str(self.test_procedure_tabel.item(i, 5).text()) if self.test_procedure_tabel.item(i, 5) else '无',
                str(self.test_procedure_tabel.item(i, 6).text()) if self.test_procedure_tabel.item(i, 6) else '无',
            ])
        test_case.case_qualified_rule = str(self.estimate_rule_eidt.text())
        test_case.case_env = self.test_env_combox.currentIndex()
        test_case.case_qualified_method = self.qualified_method_combox.currentIndex()
        test_case.case_safe_secret = str(self.safe_secret_edit.text())
        test_case.test_person = self.test_person_combox.currentIndex()
        test_case.test_join_person = str(self.test_person_join_edit.text())
        test_case.test_date = str(self.test_date_timepickedit.text())
        test_case.case_data = str(self.test_data_edit.text())
        test_case.case_problem_sheet = str(self.problem_sheet_edit.text())
        test_case.case_correct_sheet = str(self.correct_sheet_edit.text())
        test_case.case_diff = str(self.test_diff_edit.text())
        return test_case

    def __load_config(self):
        """
        加载自身的配置文件
        :return:
        """
        try:
            with io.open('./data/config.json', encoding='utf-8') as f:
                dic = json.load(f, encoding='utf-8')
                return dic
        except Exception as e:
            QMessageBox.warning(None, u'提示', u'加载配置文件失败{0}'.format(str(e)))

    def on_generate_doc(self):
        """
        导出按钮点击事件处理
        :return:
        """
        generate_dir = QFileDialog.getExistingDirectory(self, u'选择生成文档的存储目录', '.')
        if not generate_dir:
            return

        export_item_keys = []
        for i in range(0, self.testcase_tree.topLevelItemCount()):
            if self.testcase_tree.topLevelItem(i).checkState(0) == Qt.Checked:
                export_item_keys.append(str(self.testcase_tree.topLevelItem(i).text(0)))
        doc_writer = DocWriter(export_item_keys, self.test_cases, self.config, str(generate_dir))
        doc_writer.write_doc(export_item_keys, self.test_cases, self.process_progressbar)
        doc_writer.save_doc()
