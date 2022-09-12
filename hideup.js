//. 起動
var actx;
_common_lib();

//- GUI版ならCUI版として再起動
if ( /wscript\.exe$/i.test( WScript.FullName ) ) {
	actx.shell.Run( 'CScript.exe /Nologo ' + WScript.ScriptFullName );
	WScript.Quit();
}

//.. 各種設定
var conf = {
	pc_name	: '%COMPUTERNAME%'.env_expand().toLowerCase() ,
	url_base: "https://hide.maruo.co.jp/software/" ,
	dn_base	: '%APPDATA%\\hidemaru_updater\\'.env_expand() ,
	fn_ini	: WScript.ScriptFullName.replace( '.js', '.ini' )
};
var line80 = '__________';
line80 += line80; line80 += line80; line80 += line80; 

var ini_data = {}, section_list = [];
_main_get_inidata();

var made_dir = '';
if ( ! conf.dn_base.is_folder() ) {
	actx.fs.CreateFolder( conf.dn_base );
	made_dir = ' (新規作成)';
}

var num_ver_all = _ini_info( 'num_ver', 'all' ) || '2';
var flg_open_hist = _ini_info( 'open_hist', 'all' ).bool();
var soft_name;

//.. 表示
_msg( '秀丸アップデーター', 1 );
_msg([
	'PC名: ' + conf.pc_name ,
	'デフォルトバージョン: ' + _numver2mean( num_ver_all ) ,
	'ソフト数: ' + ( section_list.length - 1 ) ,
	'インストーラー置き場フォルダ: ' + conf.dn_base + made_dir
]);

//. メイン
//.. 調査
var info_todo = {}, flg_update = false;
_main_check_all()
_msg( '調査結果', 1 );
if ( ! flg_update ) {
	_msg( '更新するべきソフトウェアはありません' );
	_end();
}
_msg( '以下のソフトウェアに更新があります' );
for ( soft_name in info_todo ) {
	_msg( soft_name + ': ' + info_todo[ soft_name ].ver, 3 );
}

if ( 
	! _ini_info( 'no_ask', 'all' ).bool() &&
	! question( 'ダウンロード・インストールを実行しますか？（"y"で実行）' ).trim().bool()
){
	_end();
}
//.. インストール
_msg( 'インストール実行', 1 );
_main_install();
_msg( '自動起動', 1 );
_main_autorun();
_end();

//. main関数
//.. _main_get_inidata
function _main_get_inidata() {
	var obj_ini = actx.fs.OpenTextFile( conf.fn_ini );
	var key, val, line, line0, section = 'dummy'; 
	while( ! obj_ini.AtEndOfStream ) {
		line = obj_ini.ReadLine().trim();
		line0 = line.charAt(0);
		if ( line0 == '[' ) {
			section = line.substring( 1, line.indexOf( ']' ) );
			section_list.push( section );
			ini_data[ section ] = {};
		} else {
			if ( line0 == ';' || line0 == '/' ) continue;
			key = line.split( '=', 2 )[0];
			val = line.substring( key.length + 1 ).trim();
			key = key.trim();
			if ( ! key || ! val ) continue;
			ini_data[ section ][ key ] = val;
		}
	}
}

//.. _main_check_all バージョン情報調査
function _main_check_all() {
	for ( var idx in section_list ) {
		soft_name = section_list[ idx ];
		if ( soft_name == 'all' || _ini_info( 'ignore' ).bool() ) continue
		_msg( soft_name, 1 );

	//... バージョン番号取得
		var page_url = _ini_info( 'url' );
		if ( ! page_url ) {
			_msg( 'URLが未定義' );
			continue;
		}
		page_url = conf.url_base + page_url;
		var page_contents = _page_contents( page_url );
		if ( ! page_contents ) {
			_msg( 'ページを取得できませんでした: ' + page_url );
			continue;
		}
		page_contents = page_contents.replace( /<!--(.|\s)+?-->/g, '' );
		var ver_list_pre = page_contents.match( />Ver([1-9].+?)</g );
		if ( ! ver_list_pre ) {
			_msg([
				'ページからバージョン番号を取得できませんでした',
				page_url ,
				'INI設定のキー url か file の値が間違っているようです'
			]);
			continue;
		}
		ver_list_pre = ver_list_pre.sort();
		var ver_list = [], done = {};
		for ( var idx = 0; ver_list_pre[ idx ]; ++ idx ) {
			var v = ver_list_pre[ idx ].replace( />Ver|</g, '' );
			if ( done[v] ) continue;
			done[v] = true;
			ver_list.unshift( v );
		}
		_msg( '公開バージョン: ' + ver_list.join( ', ' ) );

	//... インストールするべきバージョン決定
		//- [ beta, 最新, 一つ前 ] の配列にする
		if ( ! ver_list[0].is_beta() )
			ver_list.unshift( ver_list[0] );
		while( ver_list[1].is_beta() ) {
			ver_list[1] = ver_list[2];
			ver_list[2] = ver_list[3];
		}
		if ( ver_list.length == 2 )
			ver_list.push( ver_list[1] );
		var num_ver = _ini_info( 'num_ver' ) || num_ver_all;
		var ver_install = ver_list[ num_ver - 1 ];
		_msg( 'インストールするバージョン: ' + _numver2mean( num_ver ) + ' => ' + ver_install );

	//... リンクがあるかチェック
		var url = ( ver_install == ver_list[1] ? 'bin/' : 'bin3/' )
			 + _get_filename( ver_install )
		;
		if ( ! page_contents.has( url ) ) {
			_msg( 'エラー: サイト上に該当ファイルがないようです: ' + url );
			continue;
		}

	//... チェック
		var fn_local = _filename_local( ver_install );
		var reg_name = _ini_info( 'regist' );
		if ( reg_name ) {
			var ver_reg = _ver_registory( reg_name );
			_msg( 'バージョンチェック: レジストリ => ' + ver_reg );
			if ( ver_reg == ver_install ) {
				_msg( 'インストール済み', 3 );
				continue;
			}
		} else {
			_msg( 'バージョンチェック: インストーラーファイル' );
			if ( fn_local.is_file() ) {
				_msg( 'インストール済み', 3 );
				continue;
			}
		}

		//... インストール情報収集
		info_todo[ soft_name ] = {
			ver: ver_install ,
			url: conf.url_base + url ,
			local: fn_local
		}
		flg_update = true;

		//... リリースノートページ
		var hist_url = page_url.replace(
			'.html',
			ver_install.is_beta()
				? 'hist_pre.html'
				: 'hist.html'
		);
		if ( flg_open_hist ) {
			_msg( '更新あり', 2 );
			actx.shell.run( hist_url );
		} else {
			_msg( '更新あり: ' + hist_url, 2 );
		}
	}
}

//.. _main_install インストール実行
function _main_install() {
	var files_local = _files_in_dir( conf.dn_base );
	for ( soft_name in info_todo ) {
		_msg( soft_name, 1 );

		//... 古いインストーラーファイル削除
		var reg = new RegExp(
			_ini_info( 'file' ).replace( '<ver>', '_ver.+?' )
		);
		for ( var idx = 0; files_local[ idx ]; ++ idx ) {
			if ( ! reg.test( files_local[ idx ] ) ) continue;
			_msg( '古いインストーラーファイルを削除: ' + files_local[idx] );
			actx.fs.DeleteFile( conf.dn_base + files_local[idx] );
		}
		var ver_install = info_todo[ soft_name ].ver;
		var url = info_todo[ soft_name ].url;
		
		//... ダウンロード
		_msg( 'ダウンロード開始' );
		var fn_local = _filename_local( ver_install );
		if ( ! _download_file( url, fn_local ) ) {
			_msg( '失敗: ' + url, 3 );
			continue;
		}
		_msg( '完了', 3 )
		
		//... インストール
		_msg( 'インストール開始: ' + ver_install );
		try {
			actx.shell.Run( fn_local, 1, true );
		} catch( e ) {
			_msg( 'インストル中にエラー発生: ' + e, 2 );
		}
		_msg( '完了', 3 )

		//... チェック
		var	reg_name = _ini_info( 'regist' );
		if ( reg_name ) {
			if ( _ver_registory( reg_name ) == ver_install ) {
				_msg( 'インストール成功' );
			} else {
				_msg([
					'インストールしたバージョン番号がレジストリに登録されていません' ,
					'インストールに失敗したか、INI設定のregistに問題があるかのどちらかのようです' 
				]);
			}
		}
	}
}

//.. _main_autorun 自動起動
function _main_autorun() {
	var procs = {}, p;
	for(
		var e = new Enumerator( GetObject( "winmgmts:" ).InstancesOf("Win32_Process") );
		! e.atEnd();
		e.moveNext()
	) {
		p = e.item().ExecutablePath;
		if ( p ) procs[ p.toLowerCase() ] = true;
	}
	for ( soft_name in ini_data ) {
		if ( _ini_info( 'ignore' ).bool() ) continue;
		if ( ! _ini_info( 'autorun' ).bool() ) continue;
		p = _ver_registory( _ini_info( 'regist' ), 'DisplayIcon' );
		if ( ! p ) continue;
		_msg( soft_name );
		if ( procs[ p.toLowerCase() ] ) {
			_msg( 'プロセスあり', 3 );
			continue;
		}
		_msg( '起動', 3 );
		try {
			actx.shell.run( p );
		} catch (e) {
			_msg( e );
		}
	}
}

//.. _end
function _end(){
	_msg( '終了 - 以下から選択してください', 1 )
	_msg([
		'1: インストーラーファイル置き場フォルダーを開く' ,
		'2: 秀丸アップデーターのドキュメントを開く' ,
		'3: INIファイルを開く' ,
		'その他: 終了'
	])
	var ans =  question();
	if ( ans == '1' )
		actx.shell.run( conf.dn_base );
	else if ( ans == '2' )
		actx.shell.run( 'readme.html' );
	else if ( ans == '3' )
		actx.shell.run( conf.fn_ini );
	WScript.Quit();
}

//. ユーティリティ関数
//.. _ini_info
function _ini_info( key, section_name ) {
	section_name = section_name || soft_name;
	if ( ! ini_data[ section_name ] ) {
		_msg( '間違ったINI情報へのアクセス: ' + section_name );
	}
	return ini_data[ section_name ][ key + '_' + conf.pc_name ]
		|| ini_data[ section_name ][ key ]
		|| ''
	;
}

//.. _msg
function _msg( msg, lev ) {
	if ( typeof msg == 'string' )
		msg = [ msg ];
	var pre = [ '', line80 + '\n', '\t', '\t\t', '\t\t\t' ][ lev || 2 ];
	for ( var idx in msg )
		WScript.echo( pre + msg[ idx ] );
}

//.. _page_contents
function _page_contents(url){
	try {
//		var winhttp = new ActiveXObject("WinHttp.WinHttpRequest.5.1");
		var winhttp = new ActiveXObject("MSXML2.XMLHTTP");
		winhttp.open( 'GET', url, false );
		winhttp.send();
		return winhttp.responseText;
	} catch(e) {
		_msg( 'ページ読み込みエラー' + e );
		return false;
	}
}

//.. _download_file
function _download_file( url, local ) {
	try {
		var winhttp = new ActiveXObject("MSXML2.XMLHTTP");
		winhttp.open( 'GET', url, false );
		winhttp.send();
		var obj_stream = new ActiveXObject("Adodb.Stream")
		obj_stream.Type = 1; //- binary
		obj_stream.Open();
		obj_stream.Write( winhttp.responseBody );
		obj_stream.Savetofile( local, 2 ); //- create/overwite
	} catch(e) {
		_msg( 'ダウンロードエラー: ' + e );
		return false;
	}
	return true;
}

//.. _numver2mean
function _numver2mean( num ) {
	return [ '', '開発版優先', '通常の最新版', 'ひとつ前の版優先' ][ num ];
}

//.. _get_filename
function _get_filename( str_ver ) {
	return _ini_info( 'file' ).replace(
		'<ver>',
		str_ver.replace( '.', '' ).replace( 'β', 'b' )
	);
}

//.. _filename_local
function _filename_local( str_ver ) {
	return conf.dn_base + _get_filename( '_ver' + str_ver );
}

//.. _ver_registory
function _ver_registory( reg, key ) {
	reg = reg || _ini_info( 'regist' );
	if ( ! reg ) return '';
	try {
		return actx.shell.RegRead(
			'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\'
			+ reg
			+ '\\'
			+ ( key || 'DisplayVersion' )
		).replace( /V/ig, '' ).replace( 'beta', 'β' );
	} catch( e ) {
		return '';
	}
}

//.. _files_in_dir
function _files_in_dir( dn ) {
	var ret = [];
	var files = actx.fs.GetFolder( dn ).files;
	for ( var e = new Enumerator( files ); !e.atEnd(); e.moveNext() )
		ret.push( e.item().Name )
	return ret;
}

//.. question()
function question( msg ) {
	msg && _msg( msg );
	try {
		return WScript.StdIn.ReadLine();
	} catch( e ) {
		return '';
	}
}

//. 汎用関数
function _common_lib() {
	actx = {
		fs: new ActiveXObject( "Scripting.FileSystemObject" ) ,
		shell: new ActiveXObject( "WScript.Shell" )
	};

	//.. string 拡張
	String.prototype.parent = function(){
		return actx.fs.GetParentFolderName( this );
	}

	String.prototype.basename = function(){
		return actx.fs.GetBaseName( this );
	}

	String.prototype.ext = function(){
		return actx.fs.GetExtensionName( this );
	}

	String.prototype.is_file = function(){
		return actx.fs.FileExists( this );
	}

	String.prototype.is_folder = function(){
		return actx.fs.FolderExists( this );
	}

	String.prototype.q = function( rep_to ){
		return '"' + this.replace( /"/g, rep_to || '""' )  + '"';
	}

	String.prototype.has = function( str ){
		return this.indexOf( str ) != -1;
	}

	String.prototype.rightstr = function( num ) {
		return this.substr( this.length - num );
	}

	String.prototype.trim = function() {
		return this.replace( /^\s+|\s+$/g, '' );
	}

	String.prototype.env_expand = function() {
		return actx.shell.ExpandEnvironmentStrings( this );
	}

	String.prototype.bool = function() {
		return this && ' yes y true 1 '.has( ' ' + this.toLowerCase() + ' ' );
	}

	String.prototype.is_beta = function() {
		return this.has( 'b' ) || this.has( 'β' );
	}
}
