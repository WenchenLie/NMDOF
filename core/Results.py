from pathlib import Path
import numpy as np


class Results:
    def __init__(self,
        t: np.ndarray,
        base_a: np.ndarray,
        base_v: np.ndarray,
        base_u: np.ndarray,
        base_V: np.ndarray,
        ra: np.ndarray,
        rv: np.ndarray,
        ru: np.ndarray,
        mat: np.ndarray
    ):
        self.t = t  # 时间序列
        self.base_a = base_a  # 基底绝对加速度
        self.base_v = base_v  # 基底绝对速度
        self.base_u = base_u  # 基底绝对位移
        self.base_V = base_V  # 基底绝对反力
        self.ra = ra  # 楼层相对加速度
        self.rv = rv  # 楼层相对速度
        self.ru = ru  # 楼层相对位移
        self.mat = mat  # 楼层滞回响应
        self.N = ra.shape[1]  # 楼层数目
        self.NPTS = len(t)  # 时间步数

    @property
    def aa(self):
        """计算楼层绝对加速度"""
        return self.base_a[:, np.newaxis] + self.ra
    
    @property
    def av(self):
        """计算楼层绝对速度"""
        return self.base_v[:, np.newaxis] + self.rv
    
    @property
    def au(self):
        """计算楼层绝对位移"""
        return self.base_u[:, np.newaxis] + self.ru
    
    @property
    def resu(self):
        """残余相对层间位移"""
        return self.ru[-1]
    
    @property
    def all_responses(self) -> tuple[np.ndarray, ...]:
        return self.t, self.base_a, self.base_v, self.base_u, self.base_V, self.aa, self.av, self.au, self.ra, self.rv, self.ru, self.resu, self.mat

    def get_story_hysteresis(self, story_id: int) -> np.ndarray:
        """获取楼层的滞回响应，
        story_id从1到N，返回一个NPTS行2列的数组，表示楼层story_id的滞回响应
        """
        row_idx = story_id - 1
        return self.mat[:, 2 * row_idx: 2 * row_idx + 2]
    
    def get_story_shear(self, story_id: int) -> np.ndarray:
        """获取楼层的剪力，
        story_id从1到N，返回一个NPTS行1列的数组，表示楼层story_id的楼层剪力
        """
        row_idx = story_id - 1
        return self.mat[:, 2 * row_idx + 1]
    
    @classmethod
    def from_file(cls, gm_name: str, temp_path: str | Path):
        """读取计算结果

        Args:
            gm_name (str): 地震动名
            temp_path (str | Path): 临时文件夹路径

        Returns:
            Results: 返回Results的实例
        """
        result_path = Path(temp_path) / 'temp_NLMDOF_results'
        try:
            t = np.loadtxt(result_path / f'{gm_name}_base_reaction.txt')[:, 0]  # 时间序列
            base_a = np.loadtxt(result_path / f'{gm_name}_base_acc.txt')  # 基底绝对加速度
            base_v = np.loadtxt(result_path / f'{gm_name}_base_vel.txt')  # 基底绝对速度
            base_u = np.loadtxt(result_path / f'{gm_name}_base_disp.txt')  # 基底绝对位移
            base_V = np.loadtxt(result_path / f'{gm_name}_base_reaction.txt')[:, 1]  # 基底反力
            ra = np.loadtxt(result_path / f'{gm_name}_floor_acc.txt', ndmin=2)  # 楼层相对加速度
            rv = np.loadtxt(result_path / f'{gm_name}_floor_vel.txt', ndmin=2)  # 楼层相对速度
            ru = np.loadtxt(result_path / f'{gm_name}_floor_disp.txt', ndmin=2)  # 楼层相对位移
            mat = np.loadtxt(result_path / f'{gm_name}_material.txt')  # 楼层滞回响应
        except FileNotFoundError:
            raise FileNotFoundError(f'【find_result】无法找到{gm_name}的计算结果！')
        resutls = cls(t, base_a, base_v, base_u, base_V, ra, rv, ru, mat)
        return resutls


class ModeResults:
    def __init__(self,
        T: list,
        mode: list[np.ndarray]
    ):
        self.T = T  # 模态周期
        self.mode = mode  # 振型位移

    def __call__(self, mode_id: int, normalize: bool=False) -> np.ndarray:
        """获取第mode_id个模态

        Args:
            mode_id (int): 模态序号，从1开始

        Returns:
            np.ndarray: 归一化模态位移
        """
        mode = self.mode[mode_id - 1]
        if normalize:
            mode = mode / np.max(np.abs(mode))
        return mode

    @classmethod
    def from_file(cls, n: int, temp_path: str | Path):
        """读取模态结果

        Args:
            n (int): 模态阶数
            temp_path (str | Path): 临时文件夹路径

        Returns:
            tuple[list, list]: 周期、振型
        """
        mode: list[np.ndarray] = []
        result_path = Path(temp_path) / 'temp_NLMDOF_results'
        try:
            T = np.loadtxt(result_path / f'Periods.txt', ndmin=1).tolist()
            for i in range(1, n + 1):
                mode_i = np.loadtxt(result_path / f'mode_{i}.txt', ndmin=2)
                mode_i = mode_i[0, :]
                mode.append(mode_i)
        except FileNotFoundError:
            FileNotFoundError(f'【find_mode】无法找到模态结果！')
        mode_results = cls(T, mode)
        return mode_results
