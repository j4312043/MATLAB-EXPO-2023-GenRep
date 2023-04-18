function GenRep(work_dir , pic_path)
fclose('all');

%% VBSマクロ生成
% 一時的に外部に.basで保存し、Word VBA側で外部モジュールとして読み込む
vba = GenVBAScript(pic_path);

%% メモ帳でマクロ内容を表示(debug用)
% system(strcat("notepad ",vba) );

%% マクロ実行
% ivokeを使う手もありそう

%Open an ActiveX connection to Excel
Word = actxserver('word.application');

Word.Visible = true;
Word.AutomationSecurity = 1;

wb = Word.Documents.Open( strcat(pwd , '\base.docm') );
Word.Run('LoadGenRepModule');

Word.Run('GenReport');

wb.SaveAs2(strcat(work_dir , '\Sample.docm' ));
wb.Close;

Word.Quit;
Word.delete;

end