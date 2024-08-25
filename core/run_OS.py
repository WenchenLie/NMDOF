import os
from math import pi
from typing import Literal

import numpy as np
from core import opensees as ops


def run_OS_py(
        N: int,
        m: list,
        mat_lib: list[list],
        story_mat: list[list],
        th: list,
        SF: float | int,
        dt: float,
        mode_num: int,
        has_damping: bool,
        zeta_mode: tuple[int, int],
        zeta: tuple[float, float],
        setting: list,
        path: str,
        gm_name: str,
        g: float,
        print_result=False
    ) -> tuple[Literal[0, 1], list[float], list[list]]:
    """调用openseespy求解非线性多自由度

    Args:
        N (int): 自由度数量
        m (list): 各自由度质量
        mat_lib (list[list]): 材料库
        story_mat (list[list]): 每层的控制材料编号
        th (list): 地震动
        SF (float | int): 地震动放大倍数
        dt (float): 地震动步长
        mode_num (int): 模态数量
        has_damping (bool): 是否考虑阻尼
        zeta_mode (tuple[int, int]): Rayleigh阻尼的振型选用
        zeta (tuple[float, float]): Rayleigh阻尼的阻尼比
        setting (list): 求解设置：
        path (str): 临时文件夹路径
        gm_name (str): 地震动名
        g (float): 重力加速度
        print_result (bool, optional): 是否打印结果. Defaults to False.
    """
    def myprint(*str_):
        if print_result:
            print(*str_)

    myprint('========== 分析开始 ==========')
    myprint(f'层数：{N}')
    myprint(f'质量：{m}')
    myprint(f'材料：{mat_lib}')
    myprint(f'材料指派：{story_mat}')
    myprint(f'地震动步长：{dt}')
    myprint(f'最大可选模态数：{mode_num}')
    myprint(f'是否考虑阻尼：{has_damping}')
    myprint(f'阻尼振型选用：{zeta_mode}')
    myprint(f'阻尼比：{zeta}')
    myprint(f'求解设置：{setting}')

    if mode_num >= 5:
        mode_num = 5

    ops.wipe()
    ops.model('basic', '-ndm', 2, '-ndf', 3)

    # node
    ops.node(1, 0, 0)  # base node
    ops.fix(1, 1, 1, 1)
    all_node_tags = [1]
    story_nodes = []
    for i in range(N):
        ops.node(i + 2, 0, 0, '-mass', m[i], 0, 0)
        ops.fix(i + 2, 0, 1, 1)
        story_nodes.append(i + 2)
        all_node_tags.append(i + 2)
    nodeTag = i + 3
    
    # material
    for i, mat in enumerate(mat_lib):
        ops.uniaxialMaterial(*mat)
    matTag = i + 2
    
    # element
    element_tags: list[list] = []  # 各楼层包含的单元编号 [[1, 2], [3], [4, 5], ...]
    all_element_tags = []  # 所有单元编号 [1, 2, 3, 4, 5]
    current_ele_tag = 1
    for i in range(N):
        element_tags.append([])
        for j in range(len(story_mat[i])):
            ops.element('zeroLength', current_ele_tag, i + 1, i + 2, '-mat', story_mat[i][j], '-dir', 1, '-doRayleigh', 1)
            element_tags[i].append(current_ele_tag)
            all_element_tags.append(current_ele_tag)
            current_ele_tag += 1

    # Eigen analysis
    solver = '-genBandArpack' if N > 5 else '-fullGenLapack'
    lambda_ = ops.eigen(solver, mode_num)
    omg = [i ** 0.5 for i in lambda_]
    T = [2 * pi / i for i in omg]
    myprint('')
    for i, Ti in enumerate(T):
        myprint(f'T{i + 1} = {Ti}')
    np.savetxt(f'{path}/temp_NLMDOF_results/Periods.txt', T)

    # ground motion
    ops.timeSeries('Path', 1, '-dt', dt, '-values', *th, '-factor', SF * g)
    for i in range(N):
        ops.pattern('Plain', i + 1, 1, '-fact', -m[i])  # D'Alembert's principle
        ops.load(story_nodes[i], 1, 0, 0)

    # 用于读取绝对响应的零刚度SDOF
    large_m = max(m) * 1e6
    ops.node(nodeTag, 0, 0, 0)
    ops.node(nodeTag + 1, 0, 0, 0)
    ops.fix(nodeTag, 1, 1, 1)
    ops.fix(nodeTag + 1, 0, 1, 1)
    ops.mass(nodeTag + 1, large_m, 0, 0)
    static_node = nodeTag + 1  # 静止节点
    ops.uniaxialMaterial('Elastic', matTag, 0)
    ops.element('zeroLength', current_ele_tag, nodeTag, nodeTag + 1, '-mat', matTag, '-dir', 1, '-doRayleigh', 0)
    ops.pattern('Plain', i + 2, 1, '-fact', large_m)
    ops.load(static_node, 1, 0, 0)
    nodeTag += 2
    matTag += 1
    current_ele_tag += 1

    # Rayleigh damping
    if has_damping:
        z1, z2 = zeta
        if mode_num >= 2:
            # MDOF
            w1, w2 = omg[zeta_mode[0] - 1], omg[zeta_mode[1] - 1]
            a = 2 * w1 * w2 / (w2 ** 2 - w1 ** 2) * (w2 * z1 - w1 * z2)
            b = 2 * w1 * w2 / (w2 ** 2 - w1 ** 2) * (z2 / w1 - z1 / w2)
        elif mode_num == 1:
            # SDOF
            w1 = omg[0]
            a = 0
            b = 2 * z1 / w1
        myprint('阻尼: a =', a, ' b =', b)
        # ops.rayleigh(a, 0, b, 0)
        ops.region(1, '-ele', *all_element_tags, '-rayleigh', a, 0, b, 0)
        ops.region(1, '-node', *all_node_tags, '-rayleigh', a, 0, b, 0)
    else:
        myprint('无阻尼')

    # recorder
    if not os.path.exists(f'{path}/temp_NLMDOF_results'):
        os.makedirs(f'{path}/temp_NLMDOF_results')
    # 1 base node
    ops.recorder('Node', '-file', f'{path}/temp_NLMDOF_results/{gm_name}_base_reaction.txt', '-time', '-node', 1, '-dof', 1, 'reaction')
    ops.recorder('Node', '-file', f'{path}/temp_NLMDOF_results/{gm_name}_base_acc.txt', '-node', static_node, '-dof', 1, 'accel')
    ops.recorder('Node', '-file', f'{path}/temp_NLMDOF_results/{gm_name}_base_vel.txt', '-node', static_node, '-dof', 1, 'vel')
    ops.recorder('Node', '-file', f'{path}/temp_NLMDOF_results/{gm_name}_base_disp.txt', '-node', static_node, '-dof', 1, 'disp')
    # 2 floor nodes
    floor_nodes = [2 + i for i in range(N)]
    ops.recorder('Node', '-file', f'{path}/temp_NLMDOF_results/{gm_name}_floor_acc.txt', '-node', *floor_nodes, '-dof', 1, 'accel')
    ops.recorder('Node', '-file', f'{path}/temp_NLMDOF_results/{gm_name}_floor_vel.txt', '-node', *floor_nodes, '-dof', 1, 'vel')
    ops.recorder('Node', '-file', f'{path}/temp_NLMDOF_results/{gm_name}_floor_disp.txt', '-node', *floor_nodes, '-dof', 1, 'disp')
    # 3 material hysteretic curves
    ops.recorder('Element', '-file', f'{path}/temp_NLMDOF_results/{gm_name}_material.txt', '-ele', *all_element_tags, 'material', 1, 'stressStrain')
    # 4 modal results
    for i in range(1, mode_num + 1):
        ops.recorder('Node', '-file', f'{path}/temp_NLMDOF_results/mode_{i}.txt', '-node', *floor_nodes, '-dof', 1, f'eigen {i}')

    # Time history analysis
    if setting[6]:
        ops.constraints(setting[0], setting[6], setting[7])
    else:
        ops.constraints(setting[0])
    ops.numberer(setting[1])
    ops.system(setting[2])
    ops.test(setting[3], setting[8], setting[9])
    ops.algorithm(setting[4])
    ops.integrator(setting[5], setting[10], setting[11])
    ops.analysis('Transient')

    current_time = 0
    duration = dt * (len(th) - 1)
    init_dt = dt
    factor = 1
    max_factor = setting[12]
    min_factor = setting[13]
    dt_ratio = setting[14]
    done = 0
    while True:
        if current_time >= duration:
            done = 1
            break  # analysis finished
        dt = init_dt * factor * dt_ratio
        if current_time + dt > duration:
            dt = duration - current_time
        ok = ops.analyze(1, dt)
        if ok == 0:
            # current step finished
            current_time += dt
            old_factor = factor
            factor = factor * 2
            factor = min(factor, max_factor)
            dt = init_dt * factor
            if factor != old_factor:
                myprint(f'--- Enlarge factor to {factor} ---')
        else:
            # current step did not converge
            factor = factor / 4
            if factor < min_factor:
                # analysis failed
                myprint(f'--- factor is less than the minimum allowed ({factor} < {min_factor}). ---')
                myprint(f'--- Current time: {current_time}, total time: {duration}. ---')
                myprint(f'--- The analysis did not converge. ---')
                
                break  # analysis failed
            else:
                # reduce factor
                dt = init_dt * factor
                myprint(f'Current step did not converge, reduce factor to {factor}.')
    
    ops.wipeAnalysis()
    ops.wipe()
    myprint('========== 分析结束 ==========')
    
    return done, T, element_tags


if __name__ == '__main__':
    import sys
    from pathlib import Path
    sys.path.append(Path(__file__).parent.parent.as_posix())
    import matplotlib.pyplot as plt
    N = 3
    m = [2, 1, 1]
    mat_lib = [['Steel01', 1, 3000, 1500, 0.02], ['Steel01', 2, 2000, 1000, 0.02]]
    story_mat = [[1], [2], [2]]
    th = np.loadtxt('data/ChiChi.dat')[:, 1].tolist()
    SF = 1
    dt = 0.01
    mode_num = 3  # 输入值可大于5
    has_damping = True
    zeta_mode = (1, 2)
    zeta = (0.05, 0.05)
    setting = ['Transformation', 'Plain', 'BandGeneral', 'NormUnbalance', 'Newton', 'Newmark', '', '', 1e-5, 60, 0.5, 0.25, 1, 1e-6, 1]
    path = 'temp'
    gm_name = 'ChiChi'
    g = 9800

    done, T, element_tags = run_OS_py(N, m, mat_lib, story_mat, th, SF, dt, mode_num, has_damping, zeta_mode, zeta, setting, path, gm_name, g, print_result=True)
    print('done:', done)



