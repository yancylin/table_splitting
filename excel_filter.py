import pandas as pd

class ExcelFilter:
    def __init__(self):
        self.data_df = None
        self.filter_df = None

    def load_data(self, file_path, sheet_name=None):
        """加载主数据文件"""
        try:
            if file_path.endswith(('.xlsx', '.xls')):
                if sheet_name:
                    self.data_df = pd.read_excel(file_path, sheet_name=sheet_name)
                else:
                    self.data_df = pd.read_excel(file_path)
            elif file_path.endswith('.csv'):
                self.data_df = pd.read_csv(file_path, encoding='utf-8')
            return self.data_df
        except Exception as e:
            raise ValueError(f"读取数据文件失败: {str(e)}")

    def load_filter(self, file_path, sheet_name=None):
        """加载筛选条件文件"""
        try:
            if file_path.endswith(('.xlsx', '.xls')):
                if sheet_name:
                    self.filter_df = pd.read_excel(file_path, sheet_name=sheet_name)
                else:
                    self.filter_df = pd.read_excel(file_path)
            elif file_path.endswith('.csv'):
                self.filter_df = pd.read_csv(file_path, encoding='utf-8')
            return self.filter_df
        except Exception as e:
            raise ValueError(f"读取筛选条件文件失败: {str(e)}")

    def filter_by_column(self, data_column, filter_column=None):
        """
        根据指定列进行筛选
        """
        if self.data_df is None or self.filter_df is None:
            raise ValueError("请先加载数据文件和筛选条件文件")

        # 获取筛选值
        if filter_column:
            filter_values = self.filter_df[filter_column].dropna().tolist()
        else:
            filter_values = self.filter_df.iloc[:, 0].dropna().tolist()

        # 执行筛选
        filtered_result = self.data_df[self.data_df[data_column].isin(filter_values)]
        print(f"筛选结果：{len(filtered_result)} 条")
        return filtered_result

    def export_result(self, result_df, output_file='result.xlsx'):
        """导出筛选结果"""
        try:
            if output_file.endswith('.xlsx'):
                result_df.to_excel(output_file, index=False)
            elif output_file.endswith('.csv'):
                result_df.to_csv(output_file, index=False, encoding='utf-8')
            print(f"结果已导出到: {output_file}")
            return True
        except Exception as e:
            raise ValueError(f"导出文件失败: {str(e)}")

    def sum_column(self, df, column_name):
        """
        统计指定列的金额合计
        """
        if df is None:
            raise ValueError("数据为空，请先加载数据")

        if column_name not in df.columns:
            raise ValueError(f"列 '{column_name}' 不存在于数据中")

        # 计算指定列的合计值，忽略NaN值
        total = df[column_name].sum()
        print(f"列 '{column_name}' 的合计金额为: {total}")
        return total

    def find_missing_in_filter(self, data_column, filter_column):
        """
        查找数据表中在筛选条件表中不存在的记录

        参数:
        data_column: 数据表中的列名
        filter_column: 筛选条件表中的列名

        返回:
        DataFrame: 在筛选条件表中不存在的记录
        """
        # 获取筛选条件列的所有唯一值
        filter_values = set(self.filter_df[filter_column].dropna().unique())

        # 筛选出不在筛选条件中的数据
        missing_mask = ~self.data_df[data_column].isin(filter_values)
        result_df = self.data_df[missing_mask].copy()

        return result_df
