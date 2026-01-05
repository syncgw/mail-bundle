<?php
declare(strict_types=1);

/*
 * 	Administration interface handler class
 *
 *	@package	sync*gw
 *	@subpackage	Mail handler
 *	@copyright	(c) 2008 - 2026 Florian Daeumling, Germany. All right reserved
 * 	@license 	LGPL-3.0-or-later
 */

namespace syncgw\interface\mail;

use syncgw\lib\Config;
use syncgw\lib\DataStore;
use syncgw\lib\Server;
use syncgw\lib\XML;
use syncgw\interface\DBAdmin;
use syncgw\gui\guiHandler;

class Admin implements DBAdmin {

    /**
     * 	Pointer to sustaninable handlerr
     * 	@var Admin
     */
    private $_hd = null;

    /**
     * 	Singleton instance of object
     * 	@var Admin
     */
    static private $_obj = null;

    /**
	 *  Get class instance handler
	 *
	 *  @return - Class object
	 */
	public static function getInstance(): Admin {

		if (!self::$_obj) {

            self::$_obj = new self();

			// check for roundcube interface
		    if (class_exists('syncgw\\interface\\roundcube\\Admin'))
				self::$_obj->_hd = \syncgw\interface\roundcube\Admin::getInstance();
			else
				self::$_obj->hd = \syncgw\interface\mysql\Admin::getInstance();
		}

		return self::$_obj;
	}

    /**
	 * 	Show/get installation parameter
	 */
	public function getParms(): void {

		$gui = guiHandler::getInstance();
		$cnf = Config::getInstance();

		if(!($c = $gui->getVar('IMAPHost')))
			$c = $cnf->getVar(Config::IMAP_HOST);
		$gui->putQBox('IMAP server name',
					  '<input name="IMAPHost" type="text" size="40" maxlength="250" value="'.$c.'" />',
					  'IMAP server name where your mails resides (default: "localhost").', false);
		if(!($c = $gui->getVar('IMAPPort')))
			$c = $cnf->getVar(Config::IMAP_PORT);
		$gui->putQBox('IMAP port address',
					'<input name="IMAPPort" type="text" size="5" maxlength="6" value="'.$c.'" />',
					'IMAP server port (default: 143).', false);
		$enc = [
				'None'	=> '',
				'SSL'	=> 'SSL',
				'TLS'	=> 'TLS',
		];
		if(!($c = $gui->getVar('IMAPEnc')))
			$c = $cnf->getVar(Config::IMAP_ENC);
		$f = '<select name="'.'IMAPEnc">';
		foreach ($enc as $k => $v) {

			$s = $v == $c ? 'selected="selected"' : '';
			$f .= '<option '.$s.' value="'.$v.'">'.$k.'</option>';
		}
		$f .= '</select>';
		$gui->putQBox('IMAP encryption', $f,
					'Specify encryption to use for connection (default: "None").', false);
		$yn = [
				'Yes'	=> 'Y',
				'No'	=> 'N',
		];
		if(!($c = $gui->getVar('IMAPCert')))
			$c = $cnf->getVar(Config::IMAP_CERT);
		$f = '<select name="'.'IMAPCert">';
		foreach ($yn as $k => $v) {

			$s = $v == $c ? 'selected="selected"' : '';
			$f .= '<option '.$s.' value="'.$v.'">'.$k.'</option>';
		}
		$f .= '</select>';
		$gui->putQBox('IMAP server certificate validation', $f,
					'Speficfy if <strong>sync&bull;gw</strong> should request validation of server certificate for IMAP connection.', false);

		if(!($c = $gui->getVar('SMTPHost')))
			$c = $cnf->getVar(Config::SMTP_HOST);
		$gui->putQBox('SMTP server name',
					  '<input name="SMTPHost" type="text" size="40" maxlength="250" value="'.$c.'" />',
					  'SMTP server name to use for sending mails (default: "localhost").', false);
		if(!($c = $gui->getVar('SMTPPort')))
			$c = $cnf->getVar(Config::SMTP_PORT);
		$gui->putQBox('SMTP port address',
					'<input name="SMTPPort" type="text" size="5" maxlength="6" value="'.$c.'" />',
					'SMTP server port (default: 25).', false);
		$yn = [
				'Yes'	=> 'Y',
				'No'	=> 'N',
		];

		if(($c = $gui->getVar('SMTPAuth')) === null)
			$c = $cnf->getVar(Config::SMTP_AUTH);

		$f = '<select name="SMTPAuth">';
		foreach ($yn as $k => $v) {

			$s = $v == $c ? 'selected="selected"' : '';
			$f .= '<option '.$s.' value="'.$v.'">'.$k.'</option>';
		}
		$f .= '</select>';

		$gui->putQBox('SMTP authentication', $f,
			 		   'By default <strong>sync&bull;gw</strong> use SMTP authentication. Setting this option to '.
						'<strong>No</strong> disables SMTP authentication.', false);
		$enc = [
				'None'	=> '',
				'SSL'	=> 'SSL',
				'TLS'	=> 'TLS',
		];
		if(!($c = $gui->getVar('SMTPEnc')))
			$c = $cnf->getVar(Config::SMTP_ENC);
		$f = '<select name="'.'SMTPEnc">';
		foreach ($enc as $k => $v) {
			$s = $v == $c ? 'selected="selected"' : '';
			$f .= '<option '.$s.' value="'.$v.'">'.$k.'</option>';
		}
		$f .= '</select>';
		$gui->putQBox('SMTP encryption', $f,
					'Specify encryption to use for connection (default: "None").', false);

		if(!($c = $gui->getVar('ConTimeout')))
			$c = $cnf->getVar(Config::CON_TIMEOUT);
		$gui->putQBox('Connection timeout',
					'<input name="ConTimeout" type="text" size="5" maxlength="6" value="'.$c.'" />',
					'Connection test timeout in seconds (defaults to 5 seconds)', false);
		$gui->putQBox('Login credentials for connection test',
					'<input name="MAILUsr" type="text" size="20" maxlength="64" value="'.$gui->getVar('MAILUsr').'" />',
					'Login credeentials (e-mail address) used for testing the IMAP and SMTP connection (will not be stored).', false);
		$gui->putQBox('Login password for connection test',
					  '<input name="MAILUpw" type="password" size="20" maxlength="30" value="'.$gui->getVar('MAILUpw').'" />',
					  'Password for connection test (will not be stored).', false);

		$this->_hd->getParms();
	}

	/**
	 * 	Connect to handler
	 *
	 * 	@return - true=Ok; false=Error
	 */
	public function Connect(): bool {

		$gui = guiHandler::getInstance();
		$cnf = Config::getInstance();

		// connection already established?
		if ($cnf->getVar(Config::DATABASE))
			return true;

		// swap variables
		$cnf->updVar(Config::IMAP_HOST, $gui->getVar('IMAPHost'));
		if (!$cnf->getVar(Config::IMAP_HOST)) {

			$gui->clearAjax();
			$gui->putMsg('Missing IMAP server name.', Config::CSS_ERR);
			return false;
		}
		$cnf->updVar(Config::IMAP_PORT, $gui->getVar('IMAPPort'));
		if (!$cnf->getVar(Config::IMAP_PORT)) {

			$gui->clearAjax();
			$gui->putMsg('Missing IMAP server port number.', Config::CSS_ERR);
			return false;
		}
		$cnf->updVar(Config::IMAP_ENC, $gui->getVar('IMAPEnc'));
		$cnf->updVar(Config::IMAP_CERT, $gui->getVar('IMAPCert'));

		$cnf->updVar(Config::SMTP_HOST, $gui->getVar('SMTPHost'));
		if (!$cnf->getVar(Config::SMTP_HOST)) {

			$gui->clearAjax();
			$gui->putMsg('Missing SMTP server name.', Config::CSS_ERR);
			return false;
		}
		$cnf->updVar(Config::SMTP_PORT, $gui->getVar('SMTPPort'));
		if (!$cnf->getVar(Config::SMTP_PORT)) {

			$gui->clearAjax();
			$gui->putMsg('Missing SMTP server port number.', Config::CSS_ERR);
			return false;
		}
		$cnf->updVar(Config::SMTP_AUTH, $gui->getVar('SMTPAuth'));
		$cnf->updVar(Config::SMTP_ENC, $gui->getVar('SMTPEnc'));

		$cnf->updVar(Config::CON_TIMEOUT, $gui->getVar('ConTimeout'));

		$uid = $gui->getVar('MAILUsr');
		$upw = $gui->getVar('MAILUpw');

		if ($uid && $upw) {

			$hd = Handler::getInstance();
			if (!$hd->IMAP($uid, $upw, true))
				return false;
			if (!$hd->SMTP($uid, $upw, true))
				return false;
			Server::getInstance()->shutDown();
		}

		return $this->_hd->Connect();
	}

	/**
	 * 	Disconnect from handler
	 *
	 * 	@return - true=Ok; false=Error
 	 */
	public function DisConnect(): bool {

		return $this->_hd->DisConnect();
	}

	/**
	 * 	Return list of supported data store handler
	 *
	 * 	@return - Bit map of supported data store handler
	 */
	public function SupportedHandlers(): int {

		return $this->_hd->SupportedHandlers()|DataStore::MAIL;
	}

}
