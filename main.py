import pandas as pd


class MakeBgaPinList():
    def __init__(self):
        super().__init__()

    def hook(self):

        bga_df = self.read_xlsx()

        pin_map_df = self.bga_df_maker(bga_df)

        self.save_xlsx(pin_map_df)

    def read_xlsx(self):
        file_name = "bga.xlsx"
        df = pd.read_excel(file_name)
        return df

    def bga_df_maker(self, df):

        pin_int = df.columns.to_list()      # len : 27
        pin_int.pop(0)
        pin_cha = df["Unnamed: 0"].to_list()# len : 27



        pin_number = []     # 핀넘버
        for c in pin_cha:
            for i in pin_int:
                i = str(i)
                pin_number.append(c+i)
        # print(df.loc[0])
        # print(df.loc[0].to_list())
        pin_name_all = []



        for i in range(0, len(df["1"])):    # len(pin_int) = 27
            temp_list = df.loc[i].to_list()
            temp_list.pop(0)
            pin_name_all.append(temp_list)

        pin_name = []
        for i in pin_name_all:
            pin_name.extend(i)


        mydf = pd.DataFrame({"pin number" : pin_number, "pin name" : pin_name})
        out_df = mydf.dropna(axis=0)

        return out_df


    def save_xlsx(self, df):
        df.to_excel("aaa.xlsx", index=False)


if __name__ == "__main__":
    MakeBgaPinList().hook()
