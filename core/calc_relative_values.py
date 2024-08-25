import numpy as np
from scipy.interpolate import interp1d


def calc_relative_values(t_gm, y_gm, t_result, y_result):
    # 以t_gm和y_gm为基准，求结构相对响应
    f = interp1d(t_result, y_result, kind='linear', fill_value='extrapolate')
    y_result_int = f(t_gm)
    return y_result_int - y_gm


if __name__ == "__main__":
    t_gm = np.array([0, 1, 2, 3, 4])
    y_gm = np.array([0, 1, 2, 3, 4])
    t_result = np.array([0, 0.5, 1, 2, 2.5, 3, 4])
    y_result = np.array([1, 1.5, 2, 3, 3.5, 4, 5])
    y_rela = calc_relative_values(t_gm, y_gm, t_result, y_result)
    print(y_rela)
