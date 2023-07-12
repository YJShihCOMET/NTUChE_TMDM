# Read me  
> <font size=5> 此程式碼是為了計算申請台大化工系轉系、輔系或雙主修的學生成績而用</font> 
>   
> <font size=5> 此程式主要將申請本系的轉系、輔系或雙主修的學生成績中，將其修習之微積分、普通物理與普通化學成績挑出來計算平均，並計算該學生各學期所有科目之平均成績，以供系上決定是否通過申請。</font>  
>     
> <font size=3> 作者：施彥任 (F08524004@ntu.edu.tw)</font>  
> <font size=3> 最後編輯時間：2022/7/11</font>  

# 更新日誌  
<font size=8> 2023/7/12 </font>  
> <font size=4> 1. ntuche_tmdm.py: 原本學生若重複修習同一科目名稱的課程會取舊的成績，且三科平均會將其列入一並計算，更新後會改為取新的成績，而且舊的成績不列入三科平均的計算。</font>  
> <font size=4> 2. ntuche_tmdm.py: 新增功能讓結果的檔案中列每位學生的年級。</font>  
> <font size=4> 3. 台大化工系學士班轉系輔系與雙主修成績計算GH.ipynb: 修正使用教學中沒有的資料夾圖示。</font>  

<font size=3> 最後編輯時間：2023/7/12</font> 

# 編寫環境  

| <font size=4> 語言/套件 </font> | <font size=4> 版本 </font> |  
| :--------: | :--------: |  
| <font size=4> Python </font>  | <font size=4> 3.10.2 </font>  |  
| <font size=4> JupyterLab </font>  | <font size=4> 3.2.9 </font>  |  
| <font size=4> Numpy </font>  | <font size=4> 1.22.2 </font>  |  
| <font size=4> Pandas </font>  | <font size=4> 1.4.0 </font>  |  
| <font size=4> Openpyxl </font>  | <font size=4> 3.0.9 </font>  |  

<font size=3> 最後編輯時間：2023/3/23</font> 