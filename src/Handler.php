<?php
declare(strict_types=1);

/*
 * 	Mail interface handler class
 *
 *	@package	sync*gw
 *	@subpackage	Mail
 *	@copyright	(c) 2008 - 2023 Florian Daeumling, Germany. All right reserved
 * 	@license 	LGPL-3.0-or-later
 */

namespace syncgw\interface\mail;

use syncgw\lib\Config;
use syncgw\lib\DB;
use syncgw\lib\DataStore;
use syncgw\lib\Encoding;
use syncgw\lib\ErrorHandler;
use syncgw\lib\Log;
use syncgw\lib\Server;
use syncgw\lib\Util;
use syncgw\lib\XML;
use syncgw\lib\Msg;
use syncgw\gui\guiHandler;
use syncgw\lib\Attachment;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use syncgw\interface\DBextHandler;
use syncgw\document\field\fldMailTo;
use syncgw\document\field\fldMailCc;
use syncgw\document\field\fldMailFrom;
use syncgw\document\field\fldSummary;
use syncgw\document\field\fldMailReplyTo;
use syncgw\document\field\fldCreated;
use syncgw\document\field\fldThreadTopic;
use syncgw\document\field\fldImportance;
use syncgw\document\field\fldConversationId;
use syncgw\document\field\fldMailSender;
use syncgw\document\field\fldMailBcc;
use syncgw\document\field\fldMessageId;
use syncgw\document\field\fldStatus;
use syncgw\document\field\fldBody;
use syncgw\document\field\fldAttach;
use syncgw\document\field\fldRead;
use syncgw\document\field\fldIsDraft;
use syncgw\document\field\fldInternetCPID;
use syncgw\document\field\fldContentClass;
use syncgw\document\field\fldMessageClass;
use syncgw\document\field\fldConversationIndex;
use syncgw\document\field\fldGroupName;
use syncgw\document\field\fldAttribute;
use syncgw\document\field\fldBodyType;

class Handler implements DBextHandler {

	/**
	 * 	Group record
	 *
	 *  self::GROUP 	- Group id
	 *  self::NAME 		- Name
	 *  self::LOAD		- Group loaded
	 *  self::ATTR		- fldAttribute flags
	 *
	 *  Data record
	 *
	 *  self::GROUP 	- Group id
	 *
	 **/
	const GROUP       	 = 'Group';			// record group
	const NAME        	 = 'Name';			// name of record
	const LOAD 		  	 = 'Loaded';		// group is loaded
	const ATTR 		  	 = 'Attr';			// group attributes

	const FLAGS 	  	 = 'Flags';			// for mails: \Seen, \Answered, \Flagged, \Deleted, \Draft
											// for mail-Boxes: self::BOX_FLAGS
	const SEP		  	 = 'Sep';			// mail box seperator used
	const UID   	  	 = 'Uid';			// mail box unique record id
	const PATH		  	 = 'Path';			// path to mail box

	const DUMMY 		 = 'x@x.invalid';	// dummy email (invalid -> https://tools.ietf.org/html/rfc2606)

	// message types
	const TYPES 		 = [
			TYPETEXT			=> 'TEXT',
			TYPEMULTIPART		=> 'MULTIPART',
			TYPEMESSAGE			=> 'MESSAGE',
			TYPEAPPLICATION		=> 'APPLICATION',
			TYPEAUDIO			=> 'AUDIO',
			TYPEIMAGE			=> 'IMAGE',
			TYPEVIDEO			=> 'VIDEO',
			TYPEMODEL			=> 'MODEL',
			TYPEOTHER			=> 'OHER',
			9					=> 'UNKNOWN',
	];

	// encoding conversion
	const ENCODING 		 = [
	    	ENC7BIT				=> PHPMailer::ENCODING_7BIT,
			ENC8BIT				=> PHPMailer::ENCODING_8BIT,
			ENCBINARY			=> PHPMailer::ENCODING_BINARY,
			ENCBASE64			=> PHPMailer::ENCODING_BASE64,
			ENCQUOTEDPRINTABLE	=> PHPMailer::ENCODING_QUOTED_PRINTABLE,
	];

    // priority
    const PRIO 			 = [
			1 					=> '1 (Highest)',
			2 					=> '2 (High)',
			3 					=> '3 (Normal)',
    		4 					=> '4 (Low)',
    		5 					=> '5 (Lowest)',
    ];

	// mail box flags
	const BOX_FLAGS 	 = [
		LATT_NOINFERIORS	=> 'This mail box not contains, and may not contain any "children"',
		LATT_NOSELECT 		=> 'This is only a container, not a mail box',
		LATT_MARKED 		=> 'This mail box is marked. This means that it may contain new messages since the last time it was checked',
		LATT_UNMARKED 		=> 'This mail box is not marked, does not contain new messages',
		LATT_REFERRAL 		=> 'This container has a referral to a remote mail box',
		LATT_HASCHILDREN 	=> 'This mail box has selectable inferiors',
		LATT_HASNOCHILDREN 	=> 'This mail box has no selectable inferiors',
	];

	// mapping table
   	const MAP 			 = [
    // ----------------------------------------------------------------------------------------------------------------------------------------------------------
    //  0 - n mail adresses
	// 	1 - String
    //  2 - Date
    //  3 - Priority
    //  4 - Status flags
    //  5 - Ignored header
    //  6 - Skip
    // ----------------------------------------------------------------------------------------------------------------------------------------------------------
   		'subject'					=> [ 1, fldSummary::TAG, '#Subject', ],			// Title, heading, subject
		'to'						=> [ 0, fldMailTo::TAG, 'addAddress',	], 			// Primary recipients
    	'date'						=> [ 2, fldCreated::TAG, 'MessageDate' ], 		// In Internet, the date when a message was written
   	    'message-id'				=> [ 1, fldMessageId::TAG, '#MessageID' ], 		// Unique ID of this message - not used by EAS
   		'cc'						=> [ 0, fldMailCc::TAG, 'addCC', ], 				// Secondary, informational recipients
		'from'						=> [ 0, fldMailFrom::TAG, 'setFrom', ], 			// Authors or persons taking responsibility for the message
    	'from-address'				=> [ 0, fldMailFrom::TAG, 'setFrom', ], 			// Authors or persons taking responsibility for the message
    	'reply-to'					=> [ 0, fldMailReplyTo::TAG, 'addReplyTo', ],		// indicate where the sender wants replies to go
		'replyto'					=> [ 0, fldMailReplyTo::TAG, 'addReplyTo', ],		// indicate where the sender wants replies to go
   		'thread-topic'				=> [ 1, fldThreadTopic::TAG, '', ],				// thread topic
   		'x-priority'				=> [ 3, fldImportance::TAG, '#Priority' ],		// importance
    	'thread-index'				=> [ 1, fldConversationId::TAG, '', ],			// a unique identifier for a conversation
		'x-sender'					=> [ 0, fldMailSender::TAG, '' ], 				// adresses
    	'sender'					=> [ 0, fldMailSender::TAG, '#Sender' ], 			// The person or agent submitting the message
   		'bcc'						=> [ 0, fldMailBcc::TAG, 'addBCC', ],				// Recipients not to be disclosed to other recipients

   	// ----------------------------------------------------------------------------------------------------------------------------------------------------------
   	// additional fields not used by ActiveSync but required for writing e-mails

   		'flags'						=> [ 4, fldStatus::TAG, '' ],						// status flags

   	// ----------------------------------------------------------------------------------------------------------------------------------------------------------
   	// unsupported ActiveSync fields
   	//
	//  'DisplayTo'   																	// Specifies the e-mail of the primary recipients
   	//	'MessageClass'																	// Specifies the message class of this e-mail (IPM.*)
	//	'MeetingRequest'																// Contains information about the meeting
	//  'Categories'																	// collection of user-selected categories assigned
 	//  'UmCallerID'																	// callback telephone number of the person
   	//  'UmUserNotes'																	// user notes related to an electronic voice message
	//  'UmAttDuration'																	// the duration of the most recent electronic voice mail attachment in seconds
	//  'UmAttOrder'																	// identifies the order of electronic voice mail attachments
	//  'ConversationIndex'																// a set of timestamps used by clients to generate a conversation tree view
	//  'LastVerbExecuted'																// indicates the last action, such as reply or forward, that was taken on the message
	//  'LastVerbExecutionTime'															// the date and time when the action
 	//  'ReceivedAsBcc'																	// indicates to the user that they are a blind carbon copy (Bcc) recipient on the email
 	//  'AccountId'																		// specifies a unique identifier for the account that received a message
	//  'Sent'
	// ----------------------------------------------------------------------------------------------------------------------------------------------------------

    // some fields only included for syncDS() - not part of data record

   	 	'#Body'						=> [ 6, fldBody::TAG, '', ],						// were handled by self::_decode()
      	'#Attachments'				=> [ 6, fldAttach::TAG, '', ],						// were handled by self::_decode()
   	  	'#Read' 					=> [ 6, fldRead::TAG, '', ],						// automatically set
   		'#IsDraft' 					=> [ 6, fldIsDraft::TAG, '', ],						// automatically set
   	  	'#InternetCPID'				=> [ 6, fldInternetCPID::TAG, '', ],				// The original code page ID from the MIME message
	  	'#ContentClass'				=> [ 6, fldContentClass::TAG, '', ],				// content class of the data
      	'#MessageClass'				=> [ 6, fldMessageClass::TAG, '', ],
   		'#CvIndex'					=> [ 6, fldConversationIndex::TAG, '', ],
   		'#NativeBodyType'			=> [ 6, fldBodyType::TAG, '', ],

   		'#grp_name'					=> [ 6, fldGroupName::TAG, '', ],
		'#grp_attr'					=> [ 6, fldAttribute::TAG, '', ],

   	];

	// mapping table
   	const IGNORED		 = [
   	// ----------------------------------------------------------------------------------------------------------------------------------------------------------
   	// ignored headers

   		'return-path'						=> [ 5,	], 	// Used to convey the information from the MAIL FROM envelope attribute in final delivery
 		'received'							=> [ 5, ],	// Trace of MTAs which a message has passed
 		'path'								=> [ 5, ],	// List of MTAs passed
		'dl-expansion-history-indication'	=> [ 5, ],	// Trace of distribution lists passed
		'mime-version'						=> [ 5, ],	// An indicator that this message is formatted according to the MIME standard
		'control'							=> [ 5, ],	// Special Usenet News actions only
		'also-control'						=> [ 5, ],	// Special Usenet News actions and a normal article at the same time
		'original-encoded-information-types'=> [ 5, ],	// Which body part types occur in this message
		'alternate-recipient'				=> [ 5, ],	// whether this message may be forwarded to alternate recipients
		'disclose-recipients'				=> [ 5, ],	// Whether recipients are to be told the names of other recipients of the same message
		'content-disposition'				=> [ 5, ],	// Whether a MIME body part is to be shown inline or is an attachment
		'approved'							=> [ 5, ],	// Name of the moderator of the newsgroup
		'for-handling'						=> [ 5, ],	// Primary recipients, who are requested to handle the information in this message
		'for-comment'						=> [ 5, ],	// Primary recipients, who are requested to comment on the information in this message
		'newsgroups'						=> [ 5, ],	// In Usenet News: group(s) to which this article was posted
		'apparently-to'						=> [ 5, ],	// Inserted by Sendmail when there is no "To:" recipient in the original message
		'distribution'						=> [ 5, ],	// Geographical or organizational limitation
		'fax'								=> [ 5, ],	// Fax number of the originator
		'telefax'							=> [ 5, ],	// Fax number of the originator
		'phone'								=> [ 5, ],	// Phone number of the originator
		'mail-system-version'				=> [ 5, ],	// Information about the client software of the originator
		'mailer'							=> [ 5, ],	// Information about the client software of the originator
		'originating-client'				=> [ 5, ],	// Information about the client software of the originator
		'folloup-to'						=> [ 5, ],	// Used in Usenet News
  		'errors-to'							=> [ 5, ],	// Address to which notifications are to be sent and a request to get delivery notifications
  		'return-receipt-to'					=> [ 5, ],	// Address to which notifications are to be sent and a request to get delivery notifications
  		'prevent-nondelivery-report'		=> [ 5, ],	// Whether non-delivery report is wanted at delivery error
  		'generate-delivery-report'			=> [ 5, ],	// Whether a delivery report is wanted at successful delivery
  		'content-return'					=> [ 5, ],	// Indicates whether the content of a message is to be returned wit non-delivery notifications
  		'x400-content-return'				=> [ 5, ],	// Possible future change of name for "Content-Return:"
  		'content-id'						=> [ 5, ],	// Unique ID of one body part of the content of a message
 		'content-base'						=> [ 5, ],	// Base to be used for resolving relative URIs within this content part
 		'content-location'					=> [ 5, ],	// URI with which the content of this content part might be retrievable
 		'in-reply-to'						=> [ 5, ],	// Reference to message which this message is a reply to
 		'references'						=> [ 5, ],	// Reference to other related messages
 		'see-also'							=> [ 5, ],	// References to other related articles in Usenet News
		'obsoletes'							=> [ 5, ],	// Reference to previous message being corrected and replaced
		'supersedes'						=> [ 5, ],	// Commonly used in Usenet News in similar ways to the "Obsoletes" header described above
		'article-updates'					=> [ 5, ],	// Only in Usenet News
		'article-names'						=> [ 5, ],	// Reference to specially important articles for a particular Usenet Newsgroup
		'keywords'							=> [ 5, ],	// Search keys for data base retrieval
		'comments'							=> [ 5, ],	// Comments on a message
		'content-description'				=> [ 5, ],	// Description of a particular body part of a message
		'organization'						=> [ 5, ],	// Organization to which the sender of this article belongs
		'organisation'						=> [ 5, ],	// Organization to which the sender of this article belongs
		'summary'							=> [ 5, ],	// Short text describing a longer article
		'content-identifier'				=> [ 5, ],	// A text string which identifies the content of a message
		'delivery-date'						=> [ 5, ],	// The time when a message was delivered to its recipient
		'expires'							=> [ 5, ],	// A suggested expiration date
		'expiry-date'						=> [ 5, ],	// Time at which a message loses its validity
		'reply-by'							=> [ 5, ],	// Latest time at which a reply is requested
		'priority'							=> [ 5, ],	// Can be "normal", "urgent" or "non-urgent"
		'precendence'						=> [ 5, ],	// Sometimes used as a priority value
		'importance'						=> [ 5, ],	// A hint from the originator to the recipients about how important a message is
		'sensitivity'						=> [ 5, ],	// How sensitive it is to disclose this message to other people than the specified recipients
		'incomplete-copy'					=> [ 5, ],	// Body parts are missing
		'language'							=> [ 5, ],	// Can include a code for the natural language used in a message
		'content-language'					=> [ 5, ],	// Can include a code for the natural language used in a message
		'content-length'					=> [ 5, ],	// Inserted by certain mailers to indicate the size in bytes
		'lines'								=> [ 5, ],	// Size of the message
		'conversion'						=> [ 5, ],	// The body of this message may not be converted from one character set to another
		'content-conversion'				=> [ 5, ],	// The body of this message may not be converted from one character set to another
 		'content-type'						=> [ 5, ],	// Format of content (character set etc.)
		'content-sgml-entity'				=> [ 5, ],	// Information from the SGML entity declaration
		'content-transfer-encoding'			=> [ 5, ],	// Coding method used in a MIME message body
 		'message-type'						=> [ 5, ],	// indicates that this is a delivery report gatewayed from X.400
		'encoding'							=> [ 5, ],	// a kind of content-type information
		'resent-reply-to'					=> [ 5, ],	// headers referring to the forwarding
		'resent-from'						=> [ 5, ],	// headers referring to the forwarding
		'resent-sender'						=> [ 5, ],	// headers referring to the forwarding
		'resent-from'						=> [ 5, ],	// headers referring to the forwarding
		'resent-date' 						=> [ 5, ],	// headers referring to the forwarding
		'resent-to' 						=> [ 5, ],	// headers referring to the forwarding
		'resent-cc' 						=> [ 5, ],	// headers referring to the forwarding
		'resent-bcc' 						=> [ 5, ],	// headers referring to the forwarding
		'resent-message-id'					=> [ 5, ],	// headers referring to the forwarding
		'content-md5' 						=> [ 5, ],	// Checksum of content to ensur that it has not been modified
		'xref'		 						=> [ 5, ],	// Used in Usenet News
		'fcc'		 						=> [ 5, ],	// Name of file in which a copy of this message is stored
		'auto-forwarded'					=> [ 5, ],	// Has been automatically forwarded
 		'discarded-x400-ipms-extensions'	=> [ 5, ],	// Can be used in Internet mail to indicate X.400 IPM extensions
 		'discarded-x400-mts-extensions'		=> [ 5, ],	// Can be used in Internet mail to indicate X.400 MTS extensions
		'status'	 						=> [ 5, ],	// indicate the status of delivery for this message when stored

   	// ----------------------------------------------------------------------------------------------------------------------------------------------------------
  	// unknown headers

   		'in-reply-to'						=> [ 5, ],
  		'from-name'							=> [ 5, ],
  		'bounces-to'						=> [ 5, ],
   		'accept-language'					=> [ 5, ],
   		'acceptlanguage'					=> [ 5, ],
   		'auto-submitted'					=> [ 5, ],
   		'user-agent'						=> [ 5, ],
   		'authentication-results'			=> [ 5, ],
   		'dkim-signature'					=> [ 5, ],
   		'dkim-filter'						=> [ 5, ],
   		'delivered-to'						=> [ 5, ],
   		'deferred-delivery'					=> [ 5, ],
   		'mail-followup-to'					=> [ 5, ],
   		'return-receipt-to'					=> [ 5, ],
   		'disposition-notification-to'		=> [ 5, ],
   		'received-spf'						=> [ 5, ],
  		'domainkey-signature'				=> [ 5, ],
  		'tenantheader'						=> [ 5, ],
   		'arc-seal'							=> [ 5, ],
   		'affinity'							=> [ 5, ],
   		'mail id'							=> [ 5, ],
   		'message-id'						=> [ 5, ],
   		'arc-message-signature'				=> [ 5, ],
   		'arc-authentication-results'		=> [ 5, ],
   		'amq-delivery-message-id'			=> [ 5, ],
   		'list-unsubscribe-post'				=> [ 5, ],
   		'feedback-id'						=> [ 5, ],
   		'pp-correlation-id'					=> [ 5, ],
   		'origin-messageid'					=> [ 5, ],
    	'list-unsubscribe-post'				=> [ 5, ],
    	'list-id'							=> [ 5, ],
    	'list-help'							=> [ 5, ],
    	'llist-help-link'					=> [ 5, ],
    	'list-unsubscribe'					=> [ 5, ],
    	'mkatechnicalid'					=> [ 5, ],
  		'messagemaxretry'					=> [ 5, ],
  		'messageretryperiod'				=> [ 5, ],
  		'messagewebvalidityduration'		=> [ 5, ],
  		'messagevalidityduration'			=> [ 5, ],
   		'spamdiagnosticoutput'				=> [ 5, ],
     	'spamdiagnosticmetadata'			=> [ 5, ],
    	'savedfromemail'					=> [ 5, ],
    	'require-recipient-valid-since'		=> [ 5, ],
 	  	'precedence'						=> [ 5, ],
	  	'ironport-sdr'						=> [ 5, ],
	  	'ironport-hdrordr'					=> [ 5, ],
   		'msip_labels'						=> [ 5, ],
   		'sdm-mailfrom'						=> [ 5, ],
  		'list-help-link'					=> [ 5, ],

	// ----------------------------------------------------------------------------------------------------------------------------------------------------------
	];

	/**
	 * 	Synchronization preference
	 * 	@var string
	 */
	private $_pref;

   	/**
     * 	Pointer to sustainnable handler
     * 	@var Handler
     */
    private $_hd 		 = null;

    /**
     * 	Whether sustainable handler is external handler
     * 	@var boolean
     */
    private $_ext		 = false;

	/**
     * 	IMAP connection
     *  @var resource
     */
    private $_imap		 = null;

    /**
     * 	SMTP connection
     *  @var resource
     */
    private $_smtp		 = null;

    /**
     * 	IMAP host
     *  @var string
     */
    private $_host;

	/**
	 * 	Mail box structure
	 * 	false=	Mail boxes may contain sub-folders
	 * 	true= 	All folders were INBOX childrens
	 * 	@var boolean
	 */
	private $_inbox 	 = false;

	/**
	 * 	Configuration class pointer
	 * 	@var Config
	 */
	private $_cnf;

	/**
     * 	Singleton instance of object
     * 	@var Handler
     */
    static private $_obj = null;

    /**
	 *  Get class instance handler
	 *
	 *  @return - Class object
	 */
	public static function getInstance(): Handler {

		if (!self::$_obj) {

            self::$_obj = new self();
			self::$_obj->_cnf = Config::getInstance();

			// set log message codes 20401-20500
			Log::getInstance()->setLogMsg([

					// error messages
					20401 => 'Mailbox synchronization not enabled (%s)',

					// warning messages
					20402 => 'IMAP message: %s',
					20403 => 'SMTP message: %s',
					20404 => 'Error sending mail: %s',
					20405 => 'Error creating attachment [%s]',
					20406 => 'Trying to login to %s with username (%s) and password (%s)',

					20411 => 'Error reading external %s [%s]',
					20412 => 'Error adding external %s',
					20413 => 'Error updating external %s [%s]',
					20414 => 'Error deleting external %s [%s]',
					20415 => 'Mail box \'%s\' not found',
					20416 => 'Mail box \'%s\' is read only',

					20450 => 20402,
					20451 => 20402,
					20452 => 20402,
					20453 => 20402,
					20454 => 20402,
					20455 => 20402,
					20456 => 20402,
					20457 => 20402,
					20458 => 20402,
					20459 => 20402,
					20460 => 20402,
					20461 => 20402,
					20462 => 20402,
					20463 => 20402,
					20464 => 20402,
					20465 => 20402,
			]);

			// set error filters
			ErrorHandler::filter(E_WARNING, 'mail', 'imap_');

			// data base enabled?
			if (self::$_obj->_cnf->getVar(Config::DATABASE) !== 'mail')
				return self::$_obj;;

			// check for roundcube interface
		    if (class_exists('syncgw\\interface\\roundcube\\Handler')) {

				self::$_obj->_hd = \syncgw\interface\roundcube\Handler::getInstance();
				self::$_obj->_ext = true;
		    } else
				self::$_obj->_hd = \syncgw\interface\mysql\Handler::getInstance();

			// register shutdown function
			Server::getInstance()->regShutdown(__CLASS__);
		}

		return self::$_obj;
	}

    /**
	 * 	Shutdown function
	 */
	public function delInstance(): void {

		if (!self::$_obj)
			return;

		// close connection
		if (self::$_obj->_imap)
			imap_close(self::$_obj->_imap);
		if (self::$_obj->_smtp)
			self::$_obj->_smtp->smtpClose();

		self::$_obj->_hd->delInstance();

		self::$_obj = null;
	}

    /**
	 * 	Collect information about class
	 *
	 * 	@param 	- Object to store information
     *	@param 	- true = Provide status information only (if available)
	 */
	public function getInfo(XML &$xml, bool $status): void {

		$xml->addVar('Name', 'Mail data base handler');

		$class = '\\syncgw\\interface\\mail\\Admin';
		$class = $class::getInstance();
		$class->getInfo($xml, $status);

		if ($status) {

			$xml->addVar('Opt', 'Status');
			if ($this->_hd)
				$xml->addVar('Stat', 'Enabled');
			else
				$xml->addVar('Stat', 'Disabled');
		} else {

			$xml->addVar('Opt', '<a href="https://github.com/PHPMailer/PHPMailer" target="_blank">PHPMailer</a> '.
						'framework for PHP');
			$xml->addVar('Stat', 'v'.PHPMailer::VERSION);

			$xml->addVar('Opt', '<a href="https://tools.ietf.org/html/rfc2076" target="_blank">RFC2076</a> '.
						  'Common Internet Message Headers');
			$xml->addVar('Stat', 'Implemented');
			$xml->addVar('Opt', '<a href="https://tools.ietf.org/html/rfc4021" target="_blank">RFC4021</a> '.
						  'Registration of Mail and MIME Header flds');
			$xml->addVar('Stat', 'Implemented');
			$xml->addVar('Opt', '<a href="https://tools.ietf.org/html/rfc5321" target="_blank">RFC5321</a> '.
						  'Simple Mail Transfer Protocol');
			$xml->addVar('Stat', 'Implemented');
			$xml->addVar('Opt', '<a href="https://tools.ietf.org/html/rfc5322" target="_blank">RFC5322</a> '.
						  'Internet Message Format');
			$xml->addVar('Stat', 'Implemented');
			$xml->addVar('Opt', '<a href="https://tools.ietf.org/html/rfc2060" target="_blank">RFC2060</a> '.
						  'Internet message access protocol');
			$xml->addVar('Stat', 'Implemented');
			$xml->addVar('Opt', '<a href="https://tools.ietf.org/html/rfc3501" target="_blank">RFC3501</a> '.
						  'Internet message access protocol');
			$xml->addVar('Stat', 'Implemented');
		}
	}

 	/**
	 * 	Authorize user in external data base
	 *
	 * 	@param	- User name
	 * 	@param 	- Host name
	 * 	@param	- User password
	 * 	@return - true=Ok; false=Not authorized
 	 */
	public function Authorize(string $user, string $host, string $passwd): bool {

		// build full name
		if ($host)
			$user .= '@'.$host;

		if (!self::IMAP($user, $passwd))
			return false;

		if (!self::SMTP($user, $passwd))
			return false;

		if ($this->_ext)
	        return $this->_hd->Authorize($user, $host, $passwd);
		else
			return true;
	}

 	/**
	 * 	Authorize user in IMAP server
	 *
	 * 	@param	- User name
	 * 	@param	- User password
	 * 	@param  - Show GUI messages
	 * 	@return - true=Ok; false=Not authorized
 	 */
	public function IMAP(string $user, string $passwd, bool $gui = false): bool {

		// load imap configuration parameter
		$conf = [];
		foreach ([ Config::IMAP_HOST, Config::IMAP_PORT, Config::IMAP_ENC, Config::IMAP_CERT ] as $k) {

			if (($conf[$k] = $this->_cnf->getVar($k)) === null)
				return false;
		}

		$this->_host = '{'.$conf[Config::IMAP_HOST].':'.$conf[Config::IMAP_PORT].'/imap';
		switch ($conf[Config::IMAP_ENC]) {
		case 'SSL':
			$this->_host .= '/ssl';
			break;

		case 'TLS':
			$this->_host .= '/tls';

		default:
			break;
		}
		$this->_host .= '/user='.$user.($conf[Config::IMAP_CERT] == 'Y' ? '/validate-cert' : '/novalidate-cert').'}';

		Msg::InfoMsg('Connecting to imap server "'.$this->_host.'" with user "'.$user.'" and password "'.$passwd.'"');

		if ($gui) {

			$gui = guiHandler::getInstance();
			$gui->putMsg('');
			$sec = $this->_cnf->getVar(Config::CON_TIMEOUT);
        	$gui->putMsg('Testing connecting to: '.$conf[Config::IMAP_HOST], Config::CSS_INFO);
        	foreach ( [ IMAP_OPENTIMEOUT, IMAP_READTIMEOUT, IMAP_WRITETIMEOUT, IMAP_CLOSETIMEOUT ] as $typ)
	 	       	imap_timeout($typ, $sec);
		}

		if (($this->_imap = imap_open($this->_host, $user, $passwd, 0, 1)) === false) {

            Log::getInstance()->logMsg(Log::WARN, 20406, $this->_host, $user, $passwd);
            foreach (imap_errors() as $msg) {

            	if ($gui) {

            		$gui->putMsg('Connecting "'.user.'" to "'.$this->_host.'"', Config::CSS_INFO);
            		$gui->putMsg($msg, Config::CSS_INFO);
            	} else
					Log::getInstance()->logMsg(Log::WARN, 20402, $msg);
            }
            if ($gui) {

				$gui->putMsg('');
            	$gui->putMsg('+++ IMAP connection test failed', Config::CSS_ERR);
            }

            return false;
		} elseif ($gui)
			$gui->putMsg('IMAP connection test successfully', Config::CSS_INFO);

        return true;
 	}

 	/**
	 * 	Authorize user in SMTP server
	 *
	 * 	@param	- User name
	 * 	@param	- User password
	 * 	@param  - Show GUI messages
	 * 	@return - true=Ok; false=Not authorized
 	 */
	public function SMTP(string $user, string $passwd, bool $gui = false): bool {

        $this->_smtp = new PHPMailer($this->_cnf->getVar(Config::MAILER_ERR) ? true : false);

        // set debug output
        $this->_smtp->SMTPDebug = intval($this->_cnf->getVar(Config::SMTP_DEBUG));;

        // select method to send
        $this->_smtp->isSMTP();

        $this->_smtp->Host 	   		= $this->_cnf->getVar(Config::SMTP_HOST);
        $this->_smtp->Port     		= $this->_cnf->getVar(Config::SMTP_PORT);
        $this->_smtp->Username 		= $user;
        $this->_smtp->Password 		= $passwd;
        $this->_smtp->SMTPAuth 		= $this->_cnf->getVar(Config::SMTP_AUTH) == 'Y' ? true : false;
        $this->_smtp->SMTPSecure	= strtolower($this->_cnf->getVar(Config::SMTP_ENC));

		if ($gui) {

			$gui = guiHandler::getInstance();
			$gui->putMsg('');
			$gui->putMsg('Testing connecting to: '.$this->_smtp->Host, Config::CSS_INFO);
			$this->_smtp->SMTPDebug   = SMTP::DEBUG_CLIENT;
			$this->_smtp->Timeout 	  = $this->_cnf->getvar(Config::CON_TIMEOUT);
			$this->_smtp->Debugoutput = function($str, $level) {
				$gui = guiHandler::getInstance();
				$gui->putMsg(trim($str), Config::CSS_DBG);
			};
		}

		if (!$this->_smtp->smtpConnect()) {

	        Log::getInstance()->logMsg(Log::WARN, 20406, $this->_host, $user, $passwd);
			if ($gui) {

				$gui = guiHandler::getInstance();
	       		$gui->putMsg($this->_smtp->ErrorInfo, Config::CSS_WARN);
				$gui->putMsg('');
				$gui->putMsg('+++ SMTP connection test failed', Config::CSS_ERR);
			} else
    			Log::getInstance()->logMsg(Log::WARN, 20403, $this->_smtp->ErrorInfo);

			return false;
		} elseif ($gui)
	        $gui->putMsg('SMTP connection test successfully', Config::CSS_INFO);

		return true;
	}

	/**
	 * 	Excute raw SQL query on internal data base
	 *
	 * 	@param	- SQL query string
	 * 	@return	- Result string or []; null on error
	 */
	public function SQL(string $query) {

		return $this->_hd->SQL($query);
	}

	/**
	 * 	Perform query on external data base
	 *
	 * 	@param	- Handler ID
	 * 	@param	- Query command:<fieldset>
	 * 			  DataStore::ADD 	  Add record                             $parm= XML object<br>
	 * 			  DataStore::UPD 	  Update record                          $parm= XML object<br>
	 * 			  DataStore::DEL	  Delete record or group (inc. sub-recs) $parm= GUID<br>
	 * 			  DataStore::RGID     Read single record       	             $parm= GUID<br>
	 * 			  DataStore::GRPS     Read all group records                 $parm= None<br>
	 * 			  DataStore::RIDS     Read all records in group              $parm= Group ID or '' for record in base group
	 * 	@return	- According  to input parameter<fieldset>
	 * 			  DataStore::ADD 	  New record ID or false on error<br>
	 * 			  DataStore::UPD 	  true=Ok; false=Error<br>
	 * 			  DataStore::DEL	  true=Ok; false=Error<br>
	 * 			  DataStore::RGID	  XML object; false=Error<br>
	 * 			  DataStore::GRPS	  [ "GUID" => Typ of record ]<br>
	 * 			  DataStore::RIDS     [ "GUID" => Typ of record ]
	 */
	public function Query(int $hid, int $cmd, $parm = '') {

		// is it for us?
		if (!($hid & DataStore::MAIL) || !($hid & DataStore::EXT))
			return $this->_hd->Query($hid, $cmd, $parm);

		// logged on?
		if (!$this->_imap)
			return false;

		// load records?
		if (is_null($this->_ids))
			self::_loadRecs();

		$out = true;

		// check command
		switch ($cmd) {
		case DataStore::GRPS:
			// build list of records
			$out = [];
			foreach ($this->_ids as $k => $v) {

				if (substr($k, 0, 1) == DataStore::TYP_GROUP)
					$out[$k] = substr($k, 0, 1);
			}

			if ($this->_cnf->getVar(Config::DBG_SCRIPT))
				Msg::InfoMsg($out, 'All group records');
			break;

		case DataStore::RIDS:

			// find base group?

			if ($parm == '') {
				foreach ($this->_ids as $rid => $val) {

					if ((!$this->_inbox && $rid == '') ||
						($this->_inbox && !$val[self::GROUP])) {

						$parm = $rid;
						break;
					}
				}
			}

			// group found? (never should go here)
			if (!isset($this->_ids[$parm])) {

				Log::getInstance()->logMsg(Log::WARN, 20415, $parm);
				return false;
			}

			// late load group?
			if (substr($parm, 0, 1) == DataStore::TYP_GROUP && !$this->_ids[$parm][self::LOAD])
				self::_loadRecs($parm);

			// build list of records
			$out = [];
			foreach ($this->_ids as $k => $v) {

				if ($v[self::GROUP] == $parm)
					$out[$k] = substr($k, 0, 1);
			}
			if ($this->_cnf->getVar(Config::DBG_SCRIPT) )
				Msg::InfoMsg($out, 'All record ids in group "'.$parm.'"');
			break;

		case DataStore::RGID:

			if (!$parm)
				return false;

			if (!self::_chkLoad($parm) || !($out = self::_swap2int($parm))) {

				Log::getInstance()->logMsg(Log::WARN, 20411, substr($parm, 0, 1) == DataStore::TYP_DATA ? 'mail' :
							'mail box', $parm);
				return false;
			}
			break;

		case DataStore::ADD:

			// if we have no group, we switch to default group
			if (!($gid = $parm->getVar('extGroup')) || !isset($this->_ids[$gid])) {

				if (!$this->_inbox)
					$gid = '';
				else
					// set default group = INBOX
					foreach ($this->_ids as $rid => $val) {

						if (substr($rid, 0, 1) == DataStore::TYP_GROUP && $val[self::ATTR] & fldAttribute::MBOX_IN) {

							$gid = $rid;
					       	break;
						}
					}
			    $parm->updVar('extGroup', $gid);
			}

			// no group found?
			if (!isset($this->_ids[$gid])) {

				Log::getInstance()->logMsg(Log::WARN, 20415, $gid);
				return false;
			}

			// add external record
			if (!($out = self::_add($parm))) {

				Log::getInstance()->Msg(Log::WARN, 20412, $parm->getVar('Type') == DataStore::TYP_DATA ? 'mail' :
									'mail box');
				return false;
			}
			break;

		case DataStore::UPD:

			$rid = $parm->getVar('extID');

			// be sure to check record is loaded
			if (!self::_chkLoad($rid)) {

				Log::getInstance()->logMsg(Log::WARN, 20413, substr($rid, 0, 1) == DataStore::TYP_DATA ? 'mail' :
								'mail box', $rid);
				if ($this->_cnf->getVar(Config::DBG_LEVEL) == Config::DBG_TRACE)
					Msg::ErrMsg('Update should work - please check if synchronization is turned on!');
				return false;
			}

			// does record exist?
			if (!isset($this->_ids[$rid]) ||
				// is record editable?
			   	!($this->_ids[$rid][self::ATTR] & fldAttribute::EDIT) ||
				// is group writable?^
				!($this->_ids[$this->_ids[$rid][self::GROUP]][self::ATTR] & fldAttribute::WRITE)) {

				Log::getInstance()->logMsg(Log::WARN, 20416, $rid);
				return false;
			}

			// update external record
			if (!($out = self::_upd($parm))) {

				Log::getInstance()->logMsg(Log::WARN, 20413, substr($rid, 0, 1) == DataStore::TYP_DATA ? 'mail' :
								'mail box', $rid);
				return false;
			}
    		break;

		case DataStore::DEL:

			// be sure to check record is loaded
			if (!self::_chkLoad($parm)) {

				Log::getInstance()->logMsg(Log::WARN, 20414, substr($parm, 0, 1) == DataStore::TYP_DATA ? 'mail' :
								'mail box', $parm);
				return false;
			}

			// does record exist?
			if (!isset($this->_ids[$parm]) ||
				// is record a group and is it allowed to delete?
			    (substr($parm, 0, 1) == DataStore::TYP_GROUP && !($this->_ids[$parm][self::ATTR] & fldAttribute::DEL) ||
				// is record a data records and is it allowed to delete?
			   	(substr($parm, 0, 1) == DataStore::TYP_DATA &&
			   		!($this->_ids[$this->_ids[$parm][self::GROUP]][self::ATTR] & fldAttribute::WRITE)))) {

			    Log::getInstance()->logMsg(Log::WARN, 20416, $parm);
				return false;
			}

			// delete  external record
			if (!($out = self::_del($parm))) {

				Log::getInstance()->logMsg(Log::WARN, 20414, substr($parm, 0, 1) == DataStore::TYP_DATA ? 'mail' :
									'mail box', $parm);
				return false;
			}
			break;

		default:
 	  		break;
		}

		return $out;
	}

	/**
	 * 	Get list of supported fields in external data base
	 *
	 * 	@param	- Handler ID
	 * 	@return	- [ field name ]
	 */
	public function getflds(int $hid): array {

		$rc = [];

		if ($hid & DataStore::MAIL) {

			foreach (self::MAP as $k => $v)
				if ($v[0] != 6)
					$rc[] = $v[1];
			$k; // disable Eclipse warning
		} elseif ($this->_ext)
	        return $this->_hd->getflds($hid);

	    return $rc;
	}

	/**
	 * 	Reload any cached record information in external data base
	 *
	 * 	@param	- Handler ID
	 * 	@return	- true=Ok; false=Error
	 */
	public function Refresh(int $hid): bool {

		if ($hid & DataStore::MAIL) {

			self::_loadRecs();
			return true;
		}

	    return $this->_ext ? $this->_hd->Refresh($hid) : true;
	}

	/**
	 * 	Check trace record references
	 *
	 *	@param 	- Handler ID
	 * 	@param 	- External record array [ GUID ]
	 * 	@param 	- Mapping table [HID => [ GUID => NewGUID ] ]
	 */
	public function chkTrcReferences(int $hid, array $rids, array $maps): void {

		if ($this->_ext)
	    	$this->_hd->chkTrcReferences($hid, $rids, $maps);
	}

	/**
	 * 	(Re-) load existing external records
	 *
	 *  @param 	- null= root; else <GID> to load
 	 */
	private function _loadRecs(?string $grp = null): void {

		// re-create list
		if (!$grp) {

			$this->_ids = [];

	       	// get list of mail boxes
	        if (($boxes = imap_getmailboxes($this->_imap, $this->_host, "*")) === false) {

	        	foreach (imap_errors() as $msg)
	            	Log::getInstance()->logMsg(Log::DEBUG, 20450, $msg);
	            return;
	        }
		} else
			$boxes = [ $grp => 0 ];

       	// get synchronization preferences
        $p = $this->_hd->RCube->user->get_prefs();
        $this->_pref = isset($p['syncgw']) ? $p['syncgw'] : '';
        Msg::InfoMsg('Folder to synchronize "'.$this->_pref.'"');

	    // included in synchronization?
       	if ($this->_ext &&
       		strpos($this->_pref, \syncgw\interface\roundcube\Handler::MAIL_FULL.'0'.';') === false) {

 			Log::getInstance()->logMsg(Log::ERR, 20401, $this->_pref);
 			// disable any mail box from beeing loaded
 			$this->_ids = [];
   			return;
       	}

        // set counter
        $bcnt = 0;
        $par  = '';
		$done = [];

 		// get mail box subscription status
		$subs = imap_lsub($this->_imap, $this->_host, "*");
		// be sure to add inbox
		$subs[] = $this->_host.'INBOX';
		if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex' || $this->_cnf->getVar(Config::DBG_SCRIPT) == 'DBExt')
			Msg::InfoMsg($subs, 'Subscribed folders - all other will be skipped!');


		// get special mail boxes
		$p = $this->_ext ? $this->_hd->RCube->user->get_prefs() : [];
		$sbox = [
			fldAttribute::MBOX_TRASH	=> isset($p['trash_mbox']) ? $p['trash_mbox'] : '',
			fldAttribute::MBOX_DRAFT	=> isset($p['drafts_mbox']) ? $p['drafts_mbox'] : '',
			fldAttribute::MBOX_SENT	=> isset($p['sent_mbox']) ? $p['sent_mbox'] : '',
			fldAttribute::MBOX_SPAM	=> isset($p['junk_mbox']) ? $p['junk_mbox'] : '',
		];

        // scan through all mail boxes
        $l = strlen($this->_host);
        foreach ($boxes as $gid => $box) {

	        $name  = $grp ? $this->_ids[$grp][self::NAME]  : $box->name;
	        $path  = $grp ? $this->_ids[$grp][self::PATH]  : $name;
	        $delim = $grp ? $this->_ids[$grp][self::SEP]   : $box->delimiter;
	        $attr  = $grp ? $this->_ids[$grp][self::FLAGS] : $box->attributes;

        	if ($box) {

	        	// special hack to catch double references of mail boxes
	        	if (isset($done[$name]))
	        		continue;
	        	$done[$name] = 1;

	        	// convert name
	        	if (!$grp) {

		        	$path = substr($name, $l);
	        		$name = substr(imap_mutf7_to_utf8($name), $l);
	        	}

    		    $gid = DataStore::TYP_GROUP.Util::Hash($path);

    	    	// strip off parent?
    	    	if ($n = strrpos($name, $delim))
    	    		$name = substr($name, $n + 1);

    	    	if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex' ||
    	    		$this->_cnf->getVar(Config::DBG_SCRIPT) == 'DBExt')
    		    	Msg::InfoMsg($box, 'Mail box #'.$gid.' with parent "'.$par.'"');

    		    // is box subscribed?
   		    	$f = false;
   		    	foreach ($subs as $unused => $n) {

   		    		if (stripos($n, $path) !== false) {

   		    			$f = true;
   		    			break;
   		    		}
   		    	}
	    	    $unused; // disable Eclipse warning
       			if (!$f)
       				continue;

    		    $this->_ids[$gid] = [
    	    			self::GROUP	=> $par,
    					self::NAME 	=> $name,
    	    			self::PATH 	=> $path,
    	    			self::SEP	=> $delim,
    	    			self::FLAGS => $attr,
 		       			self::ATTR	=> fldAttribute::READ|fldAttribute::WRITE|fldAttribute::getMBoxType($name),
    	    			self::LOAD	=> 0,
    	    	];

    		    // check for special RoundCube mailboxes
    		    foreach ($sbox as $a => $n) {

    		    	if (!strcmp($path, $n)) {

    		    		$this->_ids[$gid][self::ATTR] &= ~fldAttribute::MBOX_USER;
    		    		$this->_ids[$gid][self::ATTR] |= $a;
    		    		break;
    		    	}
    		    }

				if ($this->_ids[$gid][self::ATTR] & fldAttribute::MBOX_USER)
					$this->_ids[$gid][self::ATTR] |= fldAttribute::EDIT|fldAttribute::DEL;

        		// inbox is always subscribed
        		if ($this->_ids[$gid][self::ATTR] & fldAttribute::MBOX_IN) {

        			// if inbox does not allow children, we fake parent
        			if ($attr & LATT_NOINFERIORS) {

        				$this->_inbox = true;
        				Msg::InfoMsg('All sub folders located in INBOX');
	    	    		$par = $gid;
        			}
        		} else {

	    	    	// get parent if in format e.g. INBOX.Sent
	    	    	if ($n = strrpos($path, $delim))
	    	    		$this->_ids[$gid][self::GROUP] = DataStore::TYP_GROUP.Util::Hash(substr($path, 0, $n));
        		}

        		// count mail box
        		$bcnt++;

        		// need to check folder?
        		if ($attr & LATT_NOSELECT)
        			continue;
        	}

        	// load mails in group?
        	if ($grp) {

        		// group loaded?
	    		$this->_ids[$grp][self::LOAD] = 1;

 		        // open mail box folder
		        if (imap_reopen($this->_imap, $this->_host.$path, OP_READONLY) === false)
	    	    	return;

		        // get mails in box
		        if (!($cnt = imap_num_msg($this->_imap)))
	    	    	$list = [];
	        	elseif (($list = imap_fetch_overview($this->_imap, '1:'.$cnt, 0)) === false) {

	        		foreach (imap_errors() as $msg)
						Log::getInstance()->logMsg(Log::DEBUG, 20451, $msg);
	        	    return;
		        }

		        if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex' ||
		        	$this->_cnf->getVar(Config::DBG_SCRIPT) == 'DBExt')
		        	Msg::InfoMsg($list, 'Mails listed in "'.$name.'"');

				if (is_array($list)) {

			        foreach ($list as $m) {

			        	$this->_ids[DataStore::TYP_DATA.$m->uid.'#'.$gid] = [
		    	    			self::GROUP	=> $gid,
		        				self::UID	=> $m->uid,
		        				self::FLAGS	=> $this->_convFlags($m),
								self::ATTR	=> fldAttribute::READ|fldAttribute::WRITE|fldAttribute::EDIT|fldAttribute::DEL,
			        	];
			        }
				}
 			}
        }

    	if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex' ||
    		$this->_cnf->getVar(Config::DBG_SCRIPT) == 'DBExt') {

        	$ids = $this->_ids;
        	foreach ($this->_ids as $id => $unused)
        		$ids[$id][Handler::ATTR] = fldAttribute::showAttr($ids[$id][Handler::ATTR]);
        	$unused; // disable Eclipse warning
        	Msg::InfoMsg($ids, 'Record mapping table ('.count($this->_ids).')');
    	}
	}

	/**
	 * 	Check record is loadeded
	 *
	 *  @param 	- Record id to load
	 *  @return - true=Ok; false=Error
 	 */
	private function _chkLoad(string $rid): bool {

		// any GUID given?
	    if (!$rid)
	    	return false;

	    // alreay loaded?
		if (!isset($this->_ids[$rid])) {

			foreach ($this->_ids as $id => $parm) {

				if (substr($id, 0, 1) == DataStore::TYP_GROUP && !$parm[self::LOAD]) {

					// load group
					self::_loadRecs($id);

					// could we load record?
					if (isset($this->_ids[$rid]))
						return true;
				}
			}
			return false;
		}

		return true;
	}

	/**
	 * 	Get external record
	 *
	 *	@param	- External record ID
	 * 	@return - Internal document or null
	 */
	private function _swap2int(string $rid): ?XML {

		$db = DB::getInstance();

		if (substr($rid, 0, 1) == DataStore::TYP_GROUP) {

			$int = $db->mkDoc(DataStore::MAIL, [
						'GID' 					=> '',
						'Typ' 					=> DataStore::TYP_GROUP,
						'extID'					=> $rid,
						'extGroup'				=> $this->_ids[$rid][self::GROUP],
						fldGroupName::TAG		=> $this->_ids[$rid][self::NAME],
						fldAttribute::TAG		=> $this->_ids[$rid][self::ATTR],
						]);

			if ($this->_cnf->getVar(Config::DBG_SCRIPT)) {

				$int->updVar(fldAttribute::TAG, fldAttribute::showAttr($this->_ids[$rid][Handler::ATTR]));
				$int->getVar('syncgw');
	            Msg::InfoMsg($int, 'Internal record');
		        $int->updVar(fldAttribute::TAG, strval($this->_ids[$rid][Handler::ATTR]));
			}
		} else {

			// load external record
			if (!($int = self::_get($rid)))
				return null;

			if ($this->_cnf->getVar(Config::DBG_SCRIPT)) {

				$int->getVar('syncgw');
    	        Msg::InfoMsg($int, 'Internal record');
			}
		}

        return $int;
	}

	/**
	 *  Get record
	 *
	 *  @param  - Record Id
	 *  @return - Internal record or null on error
	 */
	private function _get(string $rid): ?XML {

		// open mail box
		if (!imap_reopen($this->_imap, $this->_host.$this->_ids[$this->_ids[$rid][self::GROUP]][self::PATH])) {

		    foreach (imap_errors() as $msg)
				Log::getInstance()->logMsg(Log::DEBUG, 20452, $msg);
			return null;
		}

		return self::cnv2Int($rid, imap_fetchbody($this->_imap, $this->_ids[$rid][self::UID], '0', FT_PEEK|FT_UID));
	}

	/**
	 *  Add external record
	 *
	 *  @param  - XML record
	 *  @return - New record Id or null on error
	 */
	private function _add(XML &$int): ?string {

		// find parent box
		if (!$this->_inbox)
			$par = '';
		else
			// set default group = INBOX
			foreach ($this->_ids as $rid => $val) {

				if (substr($rid, 0, 1) == DataStore::TYP_GROUP && $val[self::ATTR] & fldAttribute::MBOX_IN) {

					$parent = $rid;
			       	break;
				}
			}

		if ($int->getVar('Type') == DataStore::TYP_GROUP) {

			// do not allowe any action on INBOX mail boxes
			if (fldAttribute::getMBoxType($name = $int->getVar(fldGroupName::TAG)) & fldAttribute::MBOX_IN)
				return null;

			// if we have no group, we switch to default group
			if (!($gid = $int->getVar('extGroup')) || !isset($this->_ids[$gid]))
			   	$int->updVar('extGroup', $gid = $par);

			// activate parent mail box
		   	if (!imap_reopen($this->_imap, $path = $this->_host.$this->_ids[$gid][self::PATH])) {
   		        foreach (imap_errors() as $msg)
					Log::getInstance()->logMsg(Log::DEBUG, 20453, $msg);
		   		return null;
		   	}

  			// does parent support children?
			if ($this->_ids[$gid][self::FLAGS] & LATT_NOINFERIORS)
				$path = $this->_host;
			else
				$path .= $this->_ids[$gid][self::SEP];

			$path .= imap_utf8_to_mutf7($name);

			Msg::InfoMsg('Adding new mail box "'.$path.'" for "'.$gid.'"');

			if (!imap_createmailbox($this->_imap, $path)) {

   		        foreach (imap_errors() as $msg)
					Log::getInstance()->logMsg(Log::DEBUG, 20454, $msg);
				return null;
			}
			// subscribe mail box
			if (!imap_subscribe($this->_imap, $path)) {

		        foreach (imap_errors() as $msg)
					Log::getInstance()->logMsg(Log::DEBUG, 20455, $msg);
				return null;
			}

			// load mail box information
        	$box = imap_getmailboxes($this->_imap, $this->_host, $path);

        	$path = substr($path, strlen($this->_host));
   	    	$rid = DataStore::TYP_GROUP.Util::Hash($path);
   	    	if ($a = $int->getVar(fldAttribute::TAG))
   	    		$a = intval($a);
   	    	else
   	    		$a = fldAttribute::READ|fldAttribute::WRITE|fldAttribute::getMBoxType($name);
       		$this->_ids[$rid] = [
        			self::GROUP => $gid,
   					self::NAME 	=> $name,
	       			self::PATH	=> $path,
       				self::SEP	=> $box[0]->delimiter,
        			self::FLAGS	=> $box[0]->attributes,
		 	       	self::ATTR	=> $a,
       				self::LOAD	=> 0,
         	];

			$int->updVar('extID', $rid);

			$id = $this->_ids[$rid];
    	    $id[Handler::ATTR] = fldAttribute::showAttr($id[Handler::ATTR]);
			Msg::InfoMsg($id, 'New mapping record "'.$rid.'" ('.count($this->_ids).')');

			return $rid;
		}

		// check mail box
		if (!($gid = $int->getVar('extGroup')))
			$gid = $parent;

		// get mail message
		$mime = self::cnv2MIME($int);

		// get message flags
		$int->getVar('Data');
		if ($flags = $int->getVar(fldStatus::TAG, false))
			$flags = '\\'.str_replace(',', '\\', $flags);
		else
			$flags = '';

		// add to mail box
		if (!imap_reopen($this->_imap, $this->_host.$this->_ids[$gid][self::PATH]) ||
			!imap_append($this->_imap, $this->_host.$this->_ids[$gid][self::PATH], $mime, $flags)) {

	   	    foreach (imap_errors() as $msg)
  				Log::getInstance()->logMsg(Log::DEBUG, 20456, $msg);
			return null;
		}

		// get new message number
		$m = imap_check($this->_imap);
		Msg::InfoMsg($m, 'New message data');

		// get new message id
        if (($m = imap_fetch_overview($this->_imap, $m->Nmsgs.':'.$m->Nmsgs, 0)) === false) {

  	        foreach (imap_errors() as $msg)
				Log::getInstance()->logMsg(Log::DEBUG, 20457, $msg);
            return null;
        }
		Msg::InfoMsg($m, 'Message flags');

       	$this->_ids[$rid = DataStore::TYP_DATA.$m[0]->uid.'#'.$gid] = [
	       		self::UID	=> $m[0]->uid,
	       		self::GROUP	=> $gid,
	       		self::FLAGS	=> $this->_convFlags($m[0]),
				self::ATTR	=> fldAttribute::READ|fldAttribute::WRITE|fldAttribute::EDIT|fldAttribute::DEL,
       	];
		$id = $this->_ids[$rid];
        $id[Handler::ATTR] = fldAttribute::showAttr($id[Handler::ATTR]);
		Msg::InfoMsg($id, 'New mapping record "'.$rid.'" ('.count($this->_ids).')');

		$int->updVar('extID', $rid);
		$int->updVar('extGroup', $gid);

		return $rid;
	}

	/**
	 *  Update external record
	 *
	 *  @param  - XML record
	 *  @param	- External record
	 *  @return - true or false on error
	 */
	private function _upd(XML &$int): bool {

		$rid = $int->getVar('extID');

		// check for group
		if ($int->getVar('Type') == DataStore::TYP_GROUP) {

			// do not allow any action on special mail boxes
			if (!(fldAttribute::getMBoxType($name = $int->getVar(fldGroupName::TAG)) & fldAttribute::MBOX_USER))
				return false;

			// build path
			if ($path = strrpos($this->_ids[$rid][self::PATH], $this->_ids[$rid][self::SEP]))
				$path = substr($this->_ids[$rid][self::PATH], 0, $path + 1);
			$path .= imap_utf8_to_mutf7($name);

	        // rename mail box
	        if (!imap_renamemailbox($this->_imap, $this->_host.$this->_ids[$rid][self::PATH], $this->_host.$path)) {

	        	Msg::InfoMsg('Error renaming "'.$this->_host.$this->_ids[$rid][self::PATH].
	        						'" to "'.$this->_host.$path.'"');
    		    foreach (imap_errors() as $msg)
					Log::getInstance()->logMsg(Log::DEBUG, 20458, $msg);
	        	return false;
	        }

	        // subscribe mail box
	        if (!imap_subscribe($this->_imap, $this->_host.$path)) {

	        	Msg::InfoMsg('Error subscribing "'.$this->_host.$this->_ids[$rid][self::PATH].'" to "'.
	        						$this->_host.$path.'"');
  			    foreach (imap_errors() as $msg)
					Log::getInstance()->logMsg(Log::DEBUG, 20459, $msg);
	        	return false;
	        }

			// create new record id
		    $gid = DataStore::TYP_GROUP.Util::Hash($path);
		    $int->updVar('extID', $gid);

	  		$this->_ids[$gid] = [
		        	self::GROUP => $this->_ids[$rid][self::GROUP],
	    			self::NAME 	=> $name,
	       			self::PATH	=> $path,
	        		self::SEP	=> $this->_ids[$rid][self::SEP],
	       			self::FLAGS => $this->_ids[$rid][self::FLAGS]|fldAttribute::getMBoxType($name),
	 	       		self::ATTR	=> fldAttribute::READ|fldAttribute::WRITE,
	  				self::LOAD	=> 0,
 	       	];

		    // delete old mail box from cache
		    unset($this->_ids[$rid]);

			return true;
		}

		// disable logging
		$stat = $this->_cnf->updVar(Config::TRACE, Config::TRACE_OFF);

		// emails were immutable, so we cannot update
		if (!($nrid = self::Query(DataStore::MAIL, DataStore::ADD, $int))) {

			$this->_cnf->updVar(Config::TRACE, $stat);
			return false;
		}

		// update external record id
		$int = &$int;
		$int->updVar('extID', $nrid);

		// delete old record
		self::Query(DataStore::MAIL, DataStore::DEL, $rid);

		$this->_cnf->updVar(Config::TRACE, $stat);

		return true;
	}

	/**
	 * 	Delete record
	 *
	 * 	@param 	- Record id
	 * 	@return - true=Ok, false=Error
	 */
	private function _del(string $rid): bool {

		// check for mail box
		if ($this->_ids[$rid][self::ATTR] & fldAttribute::MBOX_TALL) {

			// delete all sub records
	        foreach ($this->_ids as $id => $r) {

	           	if ($r[self::GROUP] == $rid)
	            	if (!self::Query(DataStore::MAIL, DataStore::DEL, $id))
	   	            	return false;
	        }

	        // inbox and trash cannot be deleted
	        if ($this->_ids[$rid][self::ATTR] & fldAttribute::MBOX_IN)
	           	return false;

	        // switch to INBOX
			foreach ($this->_ids as $unused => $box)
				if (isset($box[self::ATTR]) && $box[self::ATTR] & fldAttribute::MBOX_IN)
					break;
			$unused; // disable Eclipse warning

	       	imap_reopen($this->_imap, $this->_host.$box[self::PATH]);

	   		// delete mail box itself
			if (!imap_deletemailbox($this->_imap, $this->_host.$this->_ids[$rid][self::PATH])) {

		       	Msg::InfoMsg('Error deleting "'.$this->_host.$this->_ids[$rid][self::PATH].'"');
		        foreach (imap_errors() as $msg)
					Log::getInstance()->logMsg(Log::DEBUG, 20460, $msg);
		       	return false;
			}

			// delete mail box from control structure
	        unset($this->_ids[$rid]);

	        return true;
		}

		Msg::InfoMsg('Deleting "'.$rid.'" ('.$this->_ids[$rid][self::UID].')');

     	// open mail box folder
		if (imap_reopen($this->_imap, $this->_host.$this->_ids[$this->_ids[$rid][self::GROUP]][self::PATH]) === false) {

            Msg::WarnMsg('Cannot open "'.$this->_ids[$this->_ids[$rid][self::GROUP]][self::NAME].'"');
   	        foreach (imap_errors() as $msg)
				Log::getInstance()->logMsg(Log::DEBUG, 20461, $msg);
            return false;
		}

		if (imap_delete($this->_imap, $this->_ids[$rid][self::UID].':'.$this->_ids[$rid][self::UID], FT_UID) === false) {

   	        foreach (imap_errors() as $msg)
				Log::getInstance()->logMsg(Log::DEBUG, 20462, $msg);
            return false;
		}

        // trigger action (do delete)
        if (!imap_expunge($this->_imap)) {

   	        foreach (imap_errors() as $msg)
				Log::getInstance()->logMsg(Log::DEBUG, 20463, $msg);
            return false;
        }

		unset($this->_ids[$rid]);

		return true;
	}

	/**
	 * 	Convert MIME string to internal record
	 *
	 *	@param 	- External record id
	 * 	@param	- MIME message
	 * 	@return	- Internal record or null
	 */
	public function cnv2Int(string $rid, string $mime): ?XML {

		$db  = DB::getInstance();
		$xml = $db->mkDoc(DataStore::MAIL, [
					'extID' 	=> $rid,
					'extGroup'	=> $this->_ids[$rid][self::GROUP],
		]);

		// get message id
		$mid = $this->_ids[$rid][self::UID];

		// HEADER

		// split header (must be done manually, because IMAP functions does not recognize X- header)
		$hd = [];
		foreach (explode("\r\n", $mime) as $line) {

			if (!ctype_space(substr($line, 0, 1)) && strpos($line, ': ') !== false) {

				list($k, $v) = explode(': ', $line);
				$k = strtolower($k);
			} else {

				// skip empty lines
				if (!strlen($v = trim($line)))
					continue;
				// extend line?
				if (isset($hd[$k]))
					$v = '; '.$v;
			}
			// header already set?
			if (isset($hd[$k]))
				// append header
				$hd[$k] .= $v;
			else
				$hd[$k] = $v;
		}

		if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex' || $this->_cnf->getVar(Config::DBG_SCRIPT) == 'DBExt')
			Msg::InfoMsg($hd, 'Processing header');

		// scan field mapping
		foreach (self::MAP as $k => $tag) {

			// available?
			if (!isset($hd[$k]))
				continue;

			switch ($tag[0]) {
		    //  0 - n mail adresses
			case 0:
				if ($hd[$k]) {

					foreach ($this->_smtp->parseAddresses($hd[$k], false) as $v)
						$xml->addVar($tag[1], $v['name'] ? '"'.imap_utf8($v['name']).'"<'.$v['address'].'>' : $v['address']);
				}
				unset($hd[$k]);
				break;

			// 	1 - String
			case 1:
				$xml->addVar($tag[1], imap_utf8($hd[$k]));
				unset($hd[$k]);
				break;

		    //  2 - Date
			case 2:
				// special hack to stripp off additional parameter
				// "Thu, 9 Apr 2020 10:50:30 +0800 (UTC+8)"
				// "Mon, 19 Jul 2021 19:13:48 +0200; ; "
				$p = strpos($hd[$k], '+') + 1;
				while (is_numeric(substr($hd[$k], $p, 1)))
					$p++;
				$hd[$k] = substr($hd[$k], 0, $p);
				$xml->addVar($tag[1], Util::unxTime($hd[$k]));
				unset($hd[$k]);
				break;

		    //  3 - Priority
			case 3:
				foreach (self::PRIO as $v1 => $v2) {

					if ($hd[$k] == $v2)
						$xml->addVar($tag[1], strval($v1));
				}
				unset($hd[$k]);
				break;

			default:
				break;
			}
		}

		// specifies the content class of the data
		$xml->addVar(fldContentClass::TAG, 'urn:content-classes:message');

		// add message flags
		if ($this->_ids[$rid][self::FLAGS]) {

			// indicates the item's current status
			$xml->addVar(fldStatus::TAG, $this->_ids[$rid][self::FLAGS]);
			$xml->addVar(fldIsDraft::TAG, strpos($this->_ids[$rid][self::FLAGS], 'Draft') !== false ? '1' : '0');
			$xml->addVar(fldRead::TAG, strpos($this->_ids[$rid][self::FLAGS], 'Seen') !== false ? '1' : '0');
		} else
			$xml->addVar(fldRead::TAG, '0');

		self::_delHead($hd);

		// BODY
		$msgs = imap_fetchstructure($this->_imap, $mid, FT_UID);
		if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex' ||
			$this->_cnf->getVar(Config::DBG_SCRIPT) == 'DBExt')
			Msg::InfoMsg($msgs, 'imap_fetchstructure()');

		$body = [];
		// simple
		if (!is_bool($msgs)) {

			if (!isset($msgs->parts) || !$msgs->parts)
				self::_decode($xml, $msgs, $mid, '0', $body);
			else {

				// multipart message
				foreach ($msgs->parts as $no => $msg)
					self::_decode($xml, $msg, $mid, strval($no + 1), $body);
			}
		}

		// save all bodies
		foreach ($body as $typ => $data) {

 	        if (substr($data, -12) == '<br><br>')
	    		$data = substr($data, 0, -12);
		    else
		    	$data = trim($data);
		    if ($data)
	    		$xml->addVar(fldBody::TAG, $data, false, [ 'X-TYP' => $typ ]);
		}
		if ($typ == 1 || $typ == 2 || $typ == 3)
    		$xml->addVar(fldBodyType::TAG, strval($typ));

		// check for non-HTML body
		if (!isset($body[fldBody::TYP_TXT]) && isset($data))
    		$xml->addVar(fldBody::TAG, trim(strip_tags($data)), false, [ 'X-TYP' => fldBody::TYP_TXT ]);

		$xml->getVar('Data');
		$xp = $xml->savePos();

		// ensure code page is set
		if ($xml->getVar(fldInternetCPID::TAG) === null)
			$xml->addVar(fldInternetCPID::TAG, Encoding::getInstance()->getMSCP('UTF-8'));

		// add message class
		$xml->restorePos($xp);
		$xml->addVar(fldMessageClass::TAG, 'IPM.Note');

		// add <Importance>
		// 0 (zero)	Low importance
		// 1 Normal importance
		// 2 High importance
		$xml->restorePos($xp);
		if ($xml->getVar(fldImportance::TAG) === null)
			$xml->addVar(fldImportance::TAG, '1');

		// build <ConversationId>
		$xml->restorePos($xp);
		if (!($gid = $xml->getVar(fldConversationId::TAG)))
			$xml->addVar(fldConversationId::TAG, $gid = substr(md5(strval($xml->getVar(fldSummary::TAG))), 0, 16));

		// build <ConversationIndex>
		$xml->xpath('//Data/'.fldCreated::TAG);
		$idx = [ $xml->getItem() * 1.0e7 + 116444736000000000, $gid ];
		# Msg::InfoMsg($idx, 'Creating <ConversationIndex> for '.$tme);
		$xml->restorePos($xp);
		$xml->addVar(fldConversationIndex::TAG, self::_encodeCVI($idx));

		return $xml;
	}

	/**
	 * 	Convert internal record to MIME
	 *
	 * 	@param	- Internal document
	 * 	@return - MIME message or null
	 */
	public function cnv2MIME(XML &$int): ?string {

		// clear all reciepcients
		$this->_smtp->clearAllRecipients();
		$this->_smtp->clearAttachments();

		$enc = Encoding::getInstance();

		// scan field mapping
		foreach (self::MAP as $k => $tag) {

			// destination available?
			if (!isset($tag[2]) || !$tag[2])
				continue;

			// build record
			$int->getVar('Data');

			switch ($tag[0]) {
		    //  0 - n mail addresses
			case 0:
				$int->xpath($tag[1], false);
				while ($val = $int->getItem()) {

					if (substr($tag[2], 0, 1) != '#') {
						$func = $tag[2];
						foreach (PHPMailer::parseAddresses($val) as $v)
							$this->_smtp->$func($v['address'], $v['name']);
					} else {

						$func = substr($tag[2], 1);
						foreach (PHPMailer::parseAddresses($val) as $v)
							if ($v['name'])
								$this->_smtp->$func = '"'.$v['name'].'"<'.$v['address'].'>';
							else
								$this->_smtp->$func = $v['address'];
					}
				}
				break;

			// 	1 - String
			case 1:
				$int->xpath($tag[1], false);
				while ($val = $int->getItem()) {

					$f = substr($tag[2], 1);
					$this->_smtp->$f = $val;
				}
				break;

		    //  2 - Date
			case 2:
				$int->xpath($tag[1], false);
				if ($v = $int->getitem())
					$this->_smtp->MessageDate = gmdate('D, j M Y H:i:s O', intval($v));
				break;

		    //  3 - Priority
			case 3:
				$int->xpath($tag[1], false);
				if ($v = $int->getItem()) {
					$f = substr($tag[2], 1);
					$this->_smtp->$f = self::PRIO[$v];
				}
				break;

			default:
				break;
			}
		}
		$k; // disable Eclsipse warning

		// set dummy body
		$this->_smtp->Body = ' ';

		// get character set
		if ($cs = $int->getVar(fldInternetCPID::TAG)) {

			// special hack to catch malformatted character set
			// "iso-8859-1content-transfer-encoding"
			if (substr(strtolower($cs), 0, 10) == 'iso-8859-1' && strlen($cs) > 10)
				$cs = 'iso-8859-1';
			$enc->setEncoding($cs);
			if ($cs == 20127)
				$this->_smtp->CharSet = PHPMailer::CHARSET_ASCII;
			elseif ($cs == 28591)
				$this->_smtp->CharSet = PHPMailer::CHARSET_ISO88591;
			else {

				$cs = '';
				$this->_smtp->CharSet = PHPMailer::CHARSET_UTF8;
			}
		} else
			$this->_smtp->CharSet = PHPMailer::CHARSET_UTF8;

		// now we go for body
		$int->xpath('//Data/'.fldBody::TAG);
		while (($val = $int->getItem()) !== null) {

			switch ($int->getAttr('X-TYP')) {
			case fldBody::TYP_HTML:
				$this->_smtp->isHTML(true);
				$this->_smtp->AltBody = $this->_smtp->Body;
				$this->_smtp->Body = $val;
				break;

			default:
				if ($cs)
					$val = $enc->export($val);
				$this->_smtp->Body = $val;
				break;
			}
		}

		// add attachments
		$att  = Attachment::getInstance();
		$int->xpath('//Data/'.fldAttach::TAG);

		while ($int->getItem() !== null) {

			// <DisplayName>
			$ip   = $int->savePos();
			$name = $int->getVar(fldAttach::SUB_TAG[0], false);

			// <FileReference>
			$int->restorePos($ip);
			$val = $int->getVar(fldAttach::SUB_TAG[1], false);

			// <ContentId>
			$int->restorePos($ip);
			if ($cid = $int->getVar(fldAttach::SUB_TAG[3], false)) {

				if (!$this->_smtp->addStringEmbeddedImage($att->read($val), $cid, $name,
									PHPMailer::ENCODING_BASE64, $att->getVar('MIME'))) {
					Log::getInstance()->logMsg(Log::WARN, 20405, $name);
					return null;
				}
			} else {
				if (!$this->_smtp->addStringAttachment($att->read($val), $name,
									PHPMailer::ENCODING_BASE64, $att->getVar('MIME'))) {

					Log::getInstance()->logMsg(Log::WARN, 20405, $name);
					return null;
				}
			}

			$int->restorePos($ip);
		}

		// special check required for PHPMailer to avoid exception call
		if (!strcmp($this->_smtp->From, 'root@localhost')) {

			$this->_smtp->From = '';
			$this->_smtp->FromName = '';
		}

		// prepare a message for sending
		$this->_smtp->preSend();

		$mime = $this->_smtp->getSentMIMEMessage();
	    if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex' ||
	    	$this->_cnf->getVar(Config::DBG_SCRIPT) == 'DBExt')
			echo htmlentities($mime);

		return $mime;
	}

	/**
	 * 	Send mail
	 *
	 * 	@param	- true=Save in Sent mail box; false=Only send mail
	 * 	@param	- MIME data OR XML document
	 * 	@return	- Internal XML document or null on error
	 */
	public function sendMail(bool $save, $doc): ?XML {

		$rid = 0;

		// we need a mail box to store mail to get mail properly imported

		// find either Sent or Trash mail box
		if (is_string($doc)) {

			// load records?
			if (is_null( $this->_ids))
				self::_loadRecs();

			$typ = $save ? fldAttribute::MBOX_SENT : fldAttribute::MBOX_TRASH;
			$gid = 0;
			foreach ($this->_ids as $id => $box) {

				if (isset($box[self::ATTR]) && $box[self::ATTR] & $typ) {

					$gid = $id;
					break;
				}
			}
			// mail box found?
			if (!$gid) {

				Log::getInstance()->logMsg(Log::WARN, 20415, $typ == fldAttribute::MBOX_SENT ? 'Sent' : 'Trash');
				return null;
			}

			// add to mail box
			if (!imap_reopen($this->_imap, $this->_host.$this->_ids[$gid][self::PATH]) ||
				!imap_append($this->_imap, $this->_host.$this->_ids[$gid][self::PATH], $doc)) {

				foreach (imap_errors() as $msg)
					Log::getInstance()->logMsg(Log::DEBUG, 20464, $msg);
				return null;
			}

			// get new message number
			$m = imap_check($this->_imap);
			Msg::InfoMsg($m, 'New message data');

			// get new message id
	        if (($m = imap_fetch_overview($this->_imap, $m->Nmsgs.':'.$m->Nmsgs, 0)) === false) {

	  	        foreach (imap_errors() as $msg)
					Log::getInstance()->logMsg(Log::DEBUG, 20465, $msg);
	            return null;
	        }
			Msg::InfoMsg($m, 'Message flags');

	       	$this->_ids[$rid = DataStore::TYP_DATA.$m[0]->uid.'#'.$gid] = [
		   			self::UID	=> $m[0]->uid,
		   			self::GROUP	=> $gid,
		   			self::FLAGS	=> $this->_convFlags($m[0]),
	       	];
	       	Msg::InfoMsg($this->_ids[$rid], 'New record "'.$rid.'" created');

	       	$doc = self::cnv2Int($rid, $doc);

	       	$doc->getVar('syncgw');
	       	Msg::InfoMsg($doc, 'Internal record (not stored). External record stored in "'.
	       						fldAttribute::ATTR_TXT[$typ].'"');
		}

       	// convert internal document to MIME to get SMTP data filled
       	self::cnv2MIME($doc);

       	// do the sending
 		if (!$this->_cnf->getVar(Config::DBG_LEVEL) == Config::DBG_TRACE)

 			// actually send a message via the selected mechanism
	       	if (!$this->_smtp->postSend()) {

				Log::getInstance()->logMsg(Log::WARN, 20404, $this->_smtp->ErrorInfo);
				return null;
			}

		// add record?
		if ($rid) {

			// do we need to delete record?
			if ($typ == fldAttribute::MBOX_TRASH) {

				self::Query(DataStore::EXT|DataStore::MAIL, DataStore::DEL, $rid);
				return null;
			}

			// save references
			$doc->updVar('extID', $rid);
			$doc->updVar('extGroup' , $gid);
		}

		return $doc;
	}

	/**
	 *  Decode message
	 *
	 *  @param  - Output document
	 *  @param  - Message object
	 *  @param  - Message ID
	 *  @param  - Message part
	 *  @param 	- Body parts
	 */
	private function _decode(XML &$xml, \stdClass $msg, int $mid, string $part, &$body = []): void {

   		$typ   = strtolower($msg->subtype) == 'plain' ? fldBody::TYP_TXT : fldBody::TYP_HTML;
	    $parms = [];
		$enc   = Encoding::getInstance();
		$data  = '';

	   	// get data
		if (!($msg->type == TYPEMULTIPART && $msg->subtype == 'RELATED'))
		    $data = $part ? imap_fetchbody($this->_imap, $mid, $part, FT_UID|FT_PEEK) :
					imap_body($this->_imap, $mid, FT_UID|FT_PEEK);

	    // default text
	    // decode data
	    $parms['encoding'] = isset(self::ENCODING[$msg->encoding]) ? $msg->encoding : ENCBINARY;
	    $parms['mime']	   = strtolower(self::TYPES[$msg->type].'/'.$msg->subtype);

	    // ENC7BIT				- 7bit
	    // ENC8BIT				- 8bit
	    // ENCBINARY			- Binary
	    // ENCBASE64			- Base64
	    if ($parms['encoding'] == ENCBASE64)
	        $data = base64_decode($data);
	    // ENCQUOTEDPRINTABLE	- Quoted-Printable
	    elseif ($parms['encoding'] == ENCQUOTEDPRINTABLE)
        	$data = quoted_printable_decode($data);
	    // ENCOTHER				- Other

	    // get all parameters, like charset, filenames of attachments, etc.
    	if (isset($msg->parameters))
        	foreach ($msg->parameters as $v)
            	$parms[strtolower($v->attribute)] = $v->value;
	    if (isset($msg->dparameters))
        	foreach ($msg->dparameters as $v)
            	$parms[strtolower($v->attribute)] = $v->value;

        // check for downgrading code page
        if (!$this->_smtp->has8bitChars($data))
        	$parms['charset'] = PHPMailer::CHARSET_ASCII;

        if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex' ||
        	$this->_cnf->getVar(Config::DBG_SCRIPT) == 'DBExt')
            Msg::InfoMsg($parms, 'Decoding message "'.$mid.'" Part '.$part.', Type: '.
            					self::TYPES[$msg->type].'/'.$msg->subtype);

        // save code page?
       	if (isset($parms['charset']) && ($val = $enc->getMSCP($parms['charset']))) {

			// ensure code page is given
			if (!$xml->getVar(fldInternetCPID::TAG))
	       		$xml->addVar(fldInternetCPID::TAG, $val);

		    // convert body
       		$enc->setEncoding($parms['charset']);
            $data = $enc->import($data);
       	}

        // any part with a filename is an attachment, so an attached text file (TYPETEXT) is not mistaken as the message.
	    if (isset($parms['filename']) || isset($parms['name'])) {

    	    $name = isset($parms['filename']) ? $parms['filename'] : $parms['name'];
    	    $att  = Attachment::getInstance();

        	// filename may be encoded, so see imap_mime_header_decode()
        	$xp = $xml->savePos();
        	$xml->getVar('Data');
        	$xml->addVar(fldAttach::TAG);
        	$xml->addVar(fldAttach::SUB_TAG[0], $name);
        	$xml->addVar(fldAttach::SUB_TAG[1], $att->create($data, $parms['mime'], $parms['encoding']));
			$xml->addVar('Method', '1');
			$xml->addVar('EstimatedDataSize', $att->getVar('Size'));
			// swap inline attachment id reference
			if (isset($msg->id))
				$xml->addVar(fldAttach::SUB_TAG[3], str_replace([ '<', '>' ], [ '', '' ], $msg->id));
			$xml->restorePos($xp);

        	return;
    	}

  	 	// TEXT
	    // messages may be split in different parts because of inline attachments, so append parts together with blank row.
   		if ($msg->type == \TYPETEXT) {

    		if (strtolower($msg->subtype) == 'plain')
 	            $body[$typ] = trim($data)."\n\n";
    		elseif (substr($data, -12) != '<br><br>')
	    		$body[$typ] = trim($data).'<br><br>';
    	}
	    // EMBEDDED MESSAGE
	    // many bounce notifications embed the original message as TYPEMESSAGE, but AOL uses TYPEMULTIPART, which is not handled here.
	    // there are no PHP functions to parse embedded messages, so this just appends the raw source to the main message.
    	elseif ($msg->type == TYPEMESSAGE)
	        $body[$typ] = trim($data)."\n\n";

	    // SUBPART RECURSION
    	if (isset($msg->parts) && $msg->parts) {

    		// 1.2, 1.2.1, etc.
    		foreach ($msg->parts as $sub => $bdy)
	            self::_decode($xml, $bdy, $mid, $part.'.'.($sub + 1), $body);
    	}
	}

	/**
	 * 	Convert message flags
	 *
	 * 	@param 	- Message flags
	 *  @return - Converted flags
	 */
	private function _convFlags(\stdClass $obj): string {

		// check message flags
		$flags = '';
		if ($obj->seen)
			$flags .= 'Seen,';
		if ($obj->answered)
			$flags .= 'Answered,';
		if ($obj->flagged)
			$flags .= 'Flagged,';
		if ($obj->deleted)
			$flags .= 'Deleted,';
		if ($obj->draft)
			$flags .= 'Draft,';

		return strlen($flags) ? substr($flags, 0, -1) : '';
	}

	/**
	 *  Delete unused / ignored header
	 *
	 *  @param  - Header array
	 */
	static function _delHead(array $hd): void {

		// delete X- header
		foreach ($hd as $k => $v) {

			if (substr($k, 0, 2) != 'x-')
				$hd[$k] = $v;
			else
				unset($hd[$k]);
		}

		// delete processed header
		foreach (self::MAP as $k => $v) {

			if ($v[0] == 5)
				unset($hd[$k]);
		}

		// delete ignored header
		foreach (self::IGNORED as $k => $v)
			unset($hd[$k]);

		if (count($hd))
			self::$_obj->_msg->WarnMsg($hd, '+++ Unused header fields found');
	}

	/**
	 * 	Convert MicroSoft FILETIME to Unix seconds
	 *
	 * 	@param 	- FILETIME
	 * 	@return - Unix seconds
	 */
	private function _ft2sec(int $ft): float {

		// strtotime('1970-1-1') - strtotime('1601-1-1') = 11644473600 seconds = 116444736000000000 nanoseconds
		return ($ft - 116444736000000000) / 1.0e7;
	}

	/**
	 * 	Decode <ConversationIndex>
	 *
	 *  @param 	- String (base64 encoded)
	 *  @return - [ FILETIME (Unix time stamp), GUID, n * child data ]
	 */
	public function decodeCVI(string $idx): array {

		$rc = [];

		// https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/filetime
		//
		// PR_CONVERSATION_INDEX
		//
		// https://technical.nttsecurity.com/post/102enx6/outlook-thread-index-value-analysis
		// https://www.meridiandiscovery.com/how-to/e-mail-conversation-index-metadata-computer-forensics/
		// https://community.metaspike.com/t/thread-index-header-field/175

		if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex') {

			Msg::InfoMsg(str_repeat('-', 116));
			Msg::InfoMsg('Conversation index: '.$idx);
		}

		// convert index
		if (!($idx = base64_decode($idx, true)))
			return [];

		// The header block is composed of 22 bytes, divided into three parts:
		// - One reserved byte. Its value is 1.
		// - Five bytes for the current system time converted to the FILETIME structure format.
		// - Sixteen bytes holding a GUID, or globally unique identifier.

		// typedef struct _FILETIME {
		//   	DWORD dwLowDateTime;
		//   	DWORD dwHighDateTime;
		// } FILETIME, FAR *LPFILETIME;
		//
		// Holds an unsigned 64-bit date and time value for a file. This value represents
		// the number of 100-nanosecond units since the start of January 1, 1601.

		// get file time
		$ft = bin2hex(substr($idx, 0, 6));

		// convert to nanoseconds - add two bytes to make it 8 bytes long
		$ft = unpack('J', hex2bin($ft.'0000'));

		$rc[] = $ft = $ft[1];

		// since date only works with integer, we must compile microseconds on our own
		if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex') {

			// ensure we have UTC time
			date_default_timezone_set('UTC');
			Msg::InfoMsg('Header time stamp: '.date('Y-m-d H:i:s.', intval(self::_ft2sec($ft))).
								round((self::_ft2sec($ft) - intval(self::_ft2sec($ft))) * 1.0e7).' (UTC) '.
								decbin($ft));
		}

		// https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/guid
		//
		// typedef struct _GUID {
		//   	unsigned long Data1;
		//   	unsigned short Data2;
		//   	unsigned short Data3;
		//   	unsigned char Data4[8];
		// } GUID;
		$rc[] = bin2hex(substr($idx, 6, 16));
		if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex')
			Msg::InfoMsg('GUID: '.$rc[1].' ('.(strlen($rc[1]) / 2).' bytes)');

		// strip off header
		$idx = substr($idx, 22);

		// Each child block is composed of 5 bytes, divided as follows:
		// - Code Bit (1 bit)
		// - Time Difference (31 bits)
		// - Random Number (4 bits)
		// - Sequence Count (4 bits)

		if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex') {

			Msg::InfoMsg('Number of children: '.strlen(strval($idx)) / 5);
			$cnt = 1;
		}

		while (strlen(strval($idx))) {

			// get child block
			$blk = substr($idx, 0, 5);
			$idx = substr($idx, 5);

			// convert to binary
			$blk = base_convert(bin2hex($blk), 16, 2);
			// fill up to 40 bits = 5 bytes
			$blk = str_pad($blk, 5*8, '0', STR_PAD_LEFT);
			if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex')
				Msg::InfoMsg('Child number: '.$cnt++.' 0'.$blk);

			// Thirty one bits containing the difference between the current time and the time in the header block expressed
			// in FILETIME units. This part of the child block is produced using one of two strategies, depending on the
			// value of the first bit. If this bit is zero, ScCreateConversationIndex discards the high 15 bits and
			// the low 18 bits. If this bit is one, the function discards the high 10 bits and the low 23 bits.
			$v = substr($blk, 1, 31);
			if (intval($blk[0]) & 0x01)
				$v = str_repeat('0', 10).$v.str_repeat('0', 23);
			else
				$v = str_repeat('0', 15).$v.str_repeat('0', 18);

			// get nano offset
			$v = bindec($v);

			// time difference in nano
			$rc[] = $v;

			if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex') {

				// convert to seconds
				$t = $v / 1.0e7;
				Msg::InfoMsg('Time difference: '.date('H:i:s.', intval($t)).round(($t - intval($t)) * 1.0e7).' (UTC)');
				// sum up time
				$ft += $v;
				Msg::InfoMsg('Message date: '.date('Y-m-d H:i:s.', intval(self::_ft2sec($ft))).round((self::_ft2sec($ft) -
							intval(self::_ft2sec($ft))) * 1.0e7).' (UTC)');
			}

			// - Four bits containing a random number generated by calling the Win32 function GetTickCount.
			$rc[] = $n = base_convert(substr($blk, 32, 4), 2, 10);
			if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex')
				Msg::InfoMsg('Random number: '.$n);

	    	// - Four bits containing a sequence count that is taken from part of the random number.
			$rc[] = $n = base_convert(substr($blk, 36, 4), 2, 10);
			if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex')
				Msg::InfoMsg('Sequence count: '.$n);
			$n; // disable Eclipse warning
		}

		return $rc;
	}

	/**
	 * 	Encode <ConversationIndex>
	 *
	 *  @param 	- FILETIME array
	 *  @return - base64 encoded string
	 */
	public function _encodeCVI(array $ft): string {

		// The header block is composed of 22 bytes, divided into three parts:
		// - One reserved byte. Its value is 1.
		// - Five bytes for the current system time converted to the FILETIME structure format.
		// - Sixteen bytes holding a GUID, or globally unique identifier.

		if (!count($ft))
			return '';

		// convert to 100-nanosecond unit
		$v   = array_shift($ft);
		$v   = substr(pack('J*', $v), 0, 6);
		$idx = $v;

		if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex')
			Msg::InfoMsg('FILETIME part: '.base64_encode($idx));

		// get GUID
		$v    = hex2bin(array_shift($ft));
		$idx .= $v;

		if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex')
			Msg::InfoMsg('GUID: '.substr(base64_encode($v), 0, 16));

		// process child blocks
		// - Code Bit (1 bit)
		// - Time Difference (31 bits)
		// - Random Number (4 bits)
		// - Sequence Count (4 bits)
		while (count($ft)) {

			// child block begins with a 1 bit constant
			$blk = '0';
			if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex')
				Msg::InfoMsg('1 bit constant : '.$blk);

			// get file time offset
			$v = array_shift($ft);

			// expand to 8 bytes
			$v = str_pad(decbin($v), 64, '0', STR_PAD_LEFT);
			// use bindec($v); to get original value

			// This part of the child block is produced using one of two strategies, depending on the value of the first bit.
			// - If this bit is zero, we discards the high 15 bits and the low 18 bits.
			// - If this bit is one, discards the high 10 bits and the low 23 bits count number of low zero bits

			if (substr($v, 0, 1)) {

				$v = substr($v, 10, 31);
				if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex') {

					Msg::InfoMsg('First bit is one: discards the high 10 bits and the low 23 bits count number of low zero bits');
					Msg::InfoMsg('Time difference: '.pack('H*', base_convert($v, 2, 16)).' '.$v);
				}
			} else {
				$v = substr($v, 15, 31);
				if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex') {

					Msg::InfoMsg('First bit is zero: discards the high 15 bits and the low 18 bits');
					Msg::InfoMsg('Time difference: '.hexdec(base_convert($v, 2, 16)).' '.$v);
				}
			}
			$blk .= $v;

			$v = array_shift($ft);
			if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex')
				Msg::InfoMsg('Random number  : '.str_pad(decbin($v), 4, '0', STR_PAD_LEFT).' ('.$v.')');
			$blk .= str_pad(decbin($v), 4, '0', STR_PAD_LEFT);

			// sequence count
			$v = array_shift($ft);
			if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex')
				Msg::InfoMsg('Sequence count : '.str_pad(decbin($v), 4, '0', STR_PAD_LEFT).' ('.$v.')');
			$blk .= str_pad(decbin($v), 4, '0', STR_PAD_LEFT);

			// convert to hex.
			$blk = str_pad(base_convert($blk, 2, 16), 10, '0', STR_PAD_LEFT);

			if ($this->_cnf->getVar(Config::DBG_SCRIPT) == 'cvIndex')
				Msg::InfoMsg('Child block    : '.$blk.' 0x'.$blk);

			// add child block
			$idx .= hex2bin($blk);
		}

		return base64_encode($idx);
	}

}

