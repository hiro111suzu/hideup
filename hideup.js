//. �N��
var actx;
_common_lib();

//- GUI�łȂ�CUI�łƂ��čċN��
if ( /wscript\.exe$/i.test( WScript.FullName ) ) {
	actx.shell.Run( 'CScript.exe /Nologo ' + WScript.ScriptFullName );
	WScript.Quit();
}

//.. �e��ݒ�
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
	made_dir = ' (�V�K�쐬)';
}

var num_ver_all = _ini_info( 'num_ver', 'all' ) || '2';
var flg_open_hist = _ini_info( 'open_hist', 'all' ).bool();
var soft_name;

//.. �\��
_msg( '�G�ۃA�b�v�f�[�^�[', 1 );
_msg([
	'PC��: ' + conf.pc_name ,
	'�f�t�H���g�o�[�W����: ' + _numver2mean( num_ver_all ) ,
	'�\�t�g��: ' + ( section_list.length - 1 ) ,
	'�C���X�g�[���[�u����t�H���_: ' + conf.dn_base + made_dir
]);

//. ���C��
//.. ����
var info_todo = {}, flg_update = false;
_main_check_all()
_msg( '��������', 1 );
if ( ! flg_update ) {
	_msg( '�X�V����ׂ��\�t�g�E�F�A�͂���܂���' );
	_end();
}
_msg( '�ȉ��̃\�t�g�E�F�A�ɍX�V������܂�' );
for ( soft_name in info_todo ) {
	_msg( soft_name + ': ' + info_todo[ soft_name ].ver, 3 );
}

if ( 
	! _ini_info( 'no_ask', 'all' ).bool() &&
	! question( '�_�E�����[�h�E�C���X�g�[�������s���܂����H�i"y"�Ŏ��s�j' ).trim().bool()
){
	_end();
}
//.. �C���X�g�[��
_msg( '�C���X�g�[�����s', 1 );
_main_install();
_msg( '�����N��', 1 );
_main_autorun();
_end();

//. main�֐�
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

//.. _main_check_all �o�[�W������񒲍�
function _main_check_all() {
	for ( var idx in section_list ) {
		soft_name = section_list[ idx ];
		if ( soft_name == 'all' || _ini_info( 'ignore' ).bool() ) continue
		_msg( soft_name, 1 );

	//... �o�[�W�����ԍ��擾
		var page_url = _ini_info( 'url' );
		if ( ! page_url ) {
			_msg( 'URL������`' );
			continue;
		}
		page_url = conf.url_base + page_url;
		var page_contents = _page_contents( page_url );
		if ( ! page_contents ) {
			_msg( '�y�[�W���擾�ł��܂���ł���: ' + page_url );
			continue;
		}
		page_contents = page_contents.replace( /<!--(.|\s)+?-->/g, '' );
		var ver_list_pre = page_contents.match( />Ver([1-9].+?)</g );
		if ( ! ver_list_pre ) {
			_msg([
				'�y�[�W����o�[�W�����ԍ����擾�ł��܂���ł���',
				page_url ,
				'INI�ݒ�̃L�[ url �� file �̒l���Ԉ���Ă���悤�ł�'
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
		_msg( '���J�o�[�W����: ' + ver_list.join( ', ' ) );

	//... �C���X�g�[������ׂ��o�[�W��������
		//- [ beta, �ŐV, ��O ] �̔z��ɂ���
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
		_msg( '�C���X�g�[������o�[�W����: ' + _numver2mean( num_ver ) + ' => ' + ver_install );

	//... �����N�����邩�`�F�b�N
		var url = ( ver_install == ver_list[1] ? 'bin/' : 'bin3/' )
			 + _get_filename( ver_install )
		;
		if ( ! page_contents.has( url ) ) {
			_msg( '�G���[: �T�C�g��ɊY���t�@�C�����Ȃ��悤�ł�: ' + url );
			continue;
		}

	//... �`�F�b�N
		var fn_local = _filename_local( ver_install );
		var reg_name = _ini_info( 'regist' );
		if ( reg_name ) {
			var ver_reg = _ver_registory( reg_name );
			_msg( '�o�[�W�����`�F�b�N: ���W�X�g�� => ' + ver_reg );
			if ( ver_reg == ver_install ) {
				_msg( '�C���X�g�[���ς�', 3 );
				continue;
			}
		} else {
			_msg( '�o�[�W�����`�F�b�N: �C���X�g�[���[�t�@�C��' );
			if ( fn_local.is_file() ) {
				_msg( '�C���X�g�[���ς�', 3 );
				continue;
			}
		}

		//... �C���X�g�[�������W
		info_todo[ soft_name ] = {
			ver: ver_install ,
			url: conf.url_base + url ,
			local: fn_local
		}
		flg_update = true;

		//... �����[�X�m�[�g�y�[�W
		var hist_url = page_url.replace(
			'.html',
			ver_install.is_beta()
				? 'hist_pre.html'
				: 'hist.html'
		);
		if ( flg_open_hist ) {
			_msg( '�X�V����', 2 );
			actx.shell.run( hist_url );
		} else {
			_msg( '�X�V����: ' + hist_url, 2 );
		}
	}
}

//.. _main_install �C���X�g�[�����s
function _main_install() {
	var files_local = _files_in_dir( conf.dn_base );
	for ( soft_name in info_todo ) {
		_msg( soft_name, 1 );

		//... �Â��C���X�g�[���[�t�@�C���폜
		var reg = new RegExp(
			_ini_info( 'file' ).replace( '<ver>', '_ver.+?' )
		);
		for ( var idx = 0; files_local[ idx ]; ++ idx ) {
			if ( ! reg.test( files_local[ idx ] ) ) continue;
			_msg( '�Â��C���X�g�[���[�t�@�C�����폜: ' + files_local[idx] );
			actx.fs.DeleteFile( conf.dn_base + files_local[idx] );
		}
		var ver_install = info_todo[ soft_name ].ver;
		var url = info_todo[ soft_name ].url;
		
		//... �_�E�����[�h
		_msg( '�_�E�����[�h�J�n' );
		var fn_local = _filename_local( ver_install );
		if ( ! _download_file( url, fn_local ) ) {
			_msg( '���s: ' + url, 3 );
			continue;
		}
		_msg( '����', 3 )
		
		//... �C���X�g�[��
		_msg( '�C���X�g�[���J�n: ' + ver_install );
		try {
			actx.shell.Run( fn_local, 1, true );
		} catch( e ) {
			_msg( '�C���X�g�����ɃG���[����: ' + e, 2 );
		}
		_msg( '����', 3 )

		//... �`�F�b�N
		var	reg_name = _ini_info( 'regist' );
		if ( reg_name ) {
			if ( _ver_registory( reg_name ) == ver_install ) {
				_msg( '�C���X�g�[������' );
			} else {
				_msg([
					'�C���X�g�[�������o�[�W�����ԍ������W�X�g���ɓo�^����Ă��܂���' ,
					'�C���X�g�[���Ɏ��s�������AINI�ݒ��regist�ɖ�肪���邩�̂ǂ��炩�̂悤�ł�' 
				]);
			}
		}
	}
}

//.. _main_autorun �����N��
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
			_msg( '�v���Z�X����', 3 );
			continue;
		}
		_msg( '�N��', 3 );
		try {
			actx.shell.run( p );
		} catch (e) {
			_msg( e );
		}
	}
}

//.. _end
function _end(){
	_msg( '�I�� - �ȉ�����I�����Ă�������', 1 )
	_msg([
		'1: �C���X�g�[���[�t�@�C���u����t�H���_�[���J��' ,
		'2: �G�ۃA�b�v�f�[�^�[�̃h�L�������g���J��' ,
		'3: INI�t�@�C�����J��' ,
		'���̑�: �I��'
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

//. ���[�e�B���e�B�֐�
//.. _ini_info
function _ini_info( key, section_name ) {
	section_name = section_name || soft_name;
	if ( ! ini_data[ section_name ] ) {
		_msg( '�Ԉ����INI���ւ̃A�N�Z�X: ' + section_name );
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
		_msg( '�y�[�W�ǂݍ��݃G���[' + e );
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
		_msg( '�_�E�����[�h�G���[: ' + e );
		return false;
	}
	return true;
}

//.. _numver2mean
function _numver2mean( num ) {
	return [ '', '�J���ŗD��', '�ʏ�̍ŐV��', '�ЂƂO�̔ŗD��' ][ num ];
}

//.. _get_filename
function _get_filename( str_ver ) {
	return _ini_info( 'file' ).replace(
		'<ver>',
		str_ver.replace( '.', '' ).replace( '��', 'b' )
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
		).replace( /V/ig, '' ).replace( 'beta', '��' );
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

//. �ėp�֐�
function _common_lib() {
	actx = {
		fs: new ActiveXObject( "Scripting.FileSystemObject" ) ,
		shell: new ActiveXObject( "WScript.Shell" )
	};

	//.. string �g��
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
		return this.has( 'b' ) || this.has( '��' );
	}
}
