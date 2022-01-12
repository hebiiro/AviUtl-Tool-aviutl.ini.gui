$(function(){

function loadFile(showDialog)
{
	var fileName = 'aviutl.ini';

	if (showDialog)
	{
		// comdlg32.ocx が必要。

		var OFN_FILEMUSTEXIST            = 0x00001000;

		var dialog = new ActiveXObject('MSComDlg.CommonDialog');
		dialog.Flags = OFN_FILEMUSTEXIST;
		dialog.MaxFileSize = 256;
		dialog.Filter = 'INI Files|*.ini|All Files|*.*';
		dialog.FileName = fileName;
		dialog.CancelError = true;
		try {
			dialog.ShowOpen();
		} catch (e) {
			return;
		}

		fileName = dialog.FileName;
	}

	var html = '';

	// 全セクションの全エントリをハッシュのハッシュとして取得
	var data = getini(fileName); //=> {"セクション1": {"キー1": "値1", "キー2": "値2"}}
	for (section in data) {
		html += '<tr data-section="' + section + '">';
			html += '<td colspan="2"><h2><span class="icon">&#x72;</span>' + section + '</h2></td>';
		html += '</tr>';
		for (key in data[section]) {
			html += '<tr data-section="' + section + '" data-key="' + key + '">';
				html += '<td><div class="label"><a class="icon" href="#">&#x72;</a>' + key + '</div></td>';
				html += '<td><input class="value" type="text" value="' + data[section][key] + '"></td>';
			html += '</tr>';
		}
	}

	$('#contents').html(html);
	$('.icon').click(removeNode);
}

function saveFile()
{
	var fileName = 'aviutl.ini';

	if (true)	// ファイル保存ダイアログを使用する場合は true にする。
	{
		// comdlg32.ocx が必要。

		var OFN_OVERWRITEPROMPT          = 0x00000002;
		var OFN_HIDEREADONLY             = 0x00000004;

		var dialog = new ActiveXObject('MSComDlg.CommonDialog');
		dialog.Flags = OFN_OVERWRITEPROMPT | OFN_HIDEREADONLY;
		dialog.MaxFileSize = 256;
		dialog.Filter = 'INI Files|*.ini|All Files|*.*';
		dialog.FileName = 'aviutl.ini';
		dialog.CancelError = true;
		try {
			dialog.ShowSave();
		} catch (e) {
			return;
		}

		fileName = dialog.FileName;
	}

	putini(fileName);
}

function getini(fileName)
{
//	var data = getini(fileName); //=> {"セクション1": {"キー1": "値1", "キー2": "値2"}}

	try {
		var data = {};
		var lastSection = '';
		var fs = new ActiveXObject('Scripting.FileSystemObject');
		var file = fs.OpenTextFile(fileName, 1, false);
		while (!file.AtEndOfStream) {
			var line = file.ReadLine();
			if (line.length == 0) continue;
			var result = line.match(/^\[([^\]]*)\]/);
			if (result) {
				lastSection = result[1];
				data[lastSection] = {};
			}
			else if (lastSection.length != 0) {
				var result = line.match(/(.*)=(.*)/);
				if (result) {
					var key = result[1];
					var value = result[2];
					data[lastSection][key] = value;
				}
			}
		}
		file.Close();
		return data;
	} catch (e) {
		alert(fileName + ' の読み込みに失敗しました');
	}
}

function putini(fileName)
{
	try {
		// 空のファイルを作成する。
		var fs = new ActiveXObject('Scripting.FileSystemObject');
		var file = fs.OpenTextFile(fileName, 2, true);
		// ファイルに ini データを書き込む。
		var lastSection = '';
		$('tr').each(function(i, node) {
			var section = $(node).attr('data-section');
			var key = $(node).attr('data-key');
			var value = $(node).find('input').val();
			if (!section || !key) return;
			if (lastSection != section) {
				lastSection = section;
				file.WriteLine('[' + section + ']');
			}
			file.WriteLine(key + '=' + value);
		});
		file.Close();
	} catch (e) {
		alert(fileName + ' の書き込みに失敗しました');
	}
}

function removeNode()
{
	var parent = $(this).parents('tr');
	var section = parent.attr('data-section');
	var key = parent.attr('data-key');

	if (key) {
		parent.remove();
	}
	else {
		var selector = 'tr[data-section="' + section + '"]';
		$(selector).remove();
	}
}

$(document).ready(function()
{
	// カレントディレクトリを変更する。
	var fso = new ActiveXObject('Scripting.FileSystemObject');
	var shell = new ActiveXObject('WScript.Shell');
	var path = document.URL.replace('file://', '');
	path = fso.GetFile(path).ParentFolder.ParentFolder.Path;
	shell.currentDirectory = path;

	// aviutl.ini ファイルを読み込む。
	loadFile(false);
});

$('#load').click(function()
{
	// ファイルを読み込む。
	loadFile(true);
});

$('#save').click(function()
{
	// ファイルに保存する。
	saveFile();
});

});
