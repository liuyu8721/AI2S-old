def main(data_path="22.xlsx", save_path="result.xlsx"):
    """
    主函数
    """
    data_list = readExcel(data=data_path)
    res_list = []
    title = data_list[0]
    title += ["出发点经纬度", "目的地经纬度"]
    res_list.append(title)
    for one_list in data_list[1:]:
        start, end = one_list[11], one_list[12]
        Q = singleRequest(location=start)
        R = singleRequest(location=end)
        one_list += [Q, R]
        res_list.append(one_list)
        time.sleep(random.uniform(0.1, 0.8))
    write2Excel(res_list, save_path=save_path)