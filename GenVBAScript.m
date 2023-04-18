function vba_macro = GenVBAScript(pic_path )
%% 結合スクリプトの生成

%% 　基本仕様：文章末尾に表、段落を生成・追加する
vba_macro = strcat(pwd , '\GenReprot.bas');
macro = fopen(vba_macro,"w+",'n','Shift_JIS');

%% モジュール名指定
fprintf(macro,'Attribute VB_Name = "GenRep"\n\n');

%% 生成のメインマクロ生成
fprintf(macro ,'Sub GenReport()\n'); 
fprintf(macro ,'\tApplication.ScreenUpdating = False\n'); 

%% 目的
fprintf(macro ,'\tGenPurpose\n'); 

%% 特記事項
fprintf(macro ,'\tGenCommentInParticular\n'); 

%% 画像出力
% ダミー用画像の設定
if isempty(pic_path)
    pic_path = strcat(pwd,"\dummy.png");
end

fprintf(macro ,...
       '\nGenPic "%s",  "%s", "%s", "%s" \n ' ,...
            '一つ目の画像', pic_path, 'Fig.1-1' ,'MATLABからも追記できます');
            
%% ヘッダ設定
HeaderStr = '1234';
fprintf(macro ,'\tSetHeader "%s"\n',HeaderStr); 

fprintf(macro ,'\tApplication.ScreenUpdating = True\n'); 
fprintf(macro ,'End Sub\n\n'); 

%% マクロファイルを閉じる
fclose(macro);
end

