/*
The JavaMail4EWS project.
Copyright (C) 2011  Sebastian Just

This library is free software; you can redistribute it and/or
modify it under the terms of the GNU Lesser General Public
License as published by the Free Software Foundation; either
version 3.0 of the License, or (at your option) any later version.

This library is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
Lesser General Public License for more details.

You should have received a copy of the GNU Lesser General Public
License along with this library; if not, write to the Free Software
Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 */
package org.sourceforge.net.javamail4ews.util;

import java.io.IOException;
import java.io.InputStream;
import java.net.ConnectException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.Properties;

import javax.mail.AuthenticationFailedException;
import javax.mail.MessagingException;
import javax.mail.Session;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public final class Util {
	private static final Logger logger = LoggerFactory.getLogger("org.sourceforge.net.javamail4ews");
	private static final Properties defaults = new Properties();
	
	static {
		logger.info("JavaMail 4 EWS loaded in version {}\nUses Microsoft(R) software", getVersion());
	}

	private Util() {
	}

	@Override
	protected Object clone() throws CloneNotSupportedException {
		throw new CloneNotSupportedException();
	}

	public static String getVersion() {
		Package lPackage = Util.class.getPackage();
		return lPackage.getImplementationVersion();
	}


	public static Properties getConfiguration(Session pSession) {
		if (defaults.isEmpty()) {
			try {
				InputStream in = Thread.currentThread().getContextClassLoader()
						.getResourceAsStream("META-INF/javamail-ews-bridge.default.properties");
				defaults.load(in);
			} catch (IOException e) {
				logger.error("Error loading EWS bridge default properties", e);
			}			
		}
		Properties prop = new Properties();
		for(Object aKey : pSession.getProperties().keySet()) {
			Object aValue = pSession.getProperties().get(aKey);
			prop.put(aKey.toString(), aValue);
		}
		
		for(Object key : defaults.keySet()) {
			prop.put(key, defaults.get(key));
		}
		return prop;
	}

	public static ExchangeService getExchangeService(String host, int port, String user,
			String password, Session pSession) throws MessagingException {
		Properties props = getConfiguration(pSession);
		
		if (user == null) {
			return null;
		}
		if (password == null) {
			return null;
		}

		String version = props.getProperty(
				"org.sourceforge.net.javamail4ews.ExchangeVersion", "");
		ExchangeVersion serverVersion = null;
		if (!version.isEmpty()) {
			try {
				serverVersion = Enum.valueOf(ExchangeVersion.class, version);
			} catch (IllegalArgumentException e) {
				logger.info("Unknown version for exchange server: '" + version
						+ "' using default : no version specified");
			}
		}
		boolean enableTrace = Boolean.getBoolean(props.getProperty("org.sourceforge.net.javamail4ews.util.Util.EnableServiceTrace","false"));
		ExchangeService service = null;
		if (serverVersion != null) {
			service = new ExchangeService(serverVersion);
		} else {
		      service = new ExchangeService();
		}
		Integer connectionTimeout = getConnectionTimeout(pSession);
        Integer protocolTimeout = getProtocolTimeout(pSession);
        if(connectionTimeout != null) {
            logger.debug("setting timeout to {} using connection timeout value", connectionTimeout);
            service.setTimeout(connectionTimeout.intValue());
        }
        if(protocolTimeout != null) {
          logger.debug("setting protocol timeout to {} is ignored", protocolTimeout);
        }
		service.setTraceEnabled(enableTrace);

		ExchangeCredentials credentials = new WebCredentials(user, password);
		service.setCredentials(credentials);

		try {
			service.setUrl(new URI(host));
		} catch (URISyntaxException e) {
			throw new MessagingException(e.getMessage(), e);
		}

		try {
			//Bind to check if connection parameters are valid
			if (Boolean.getBoolean(props.getProperty("org.sourceforge.net.javamail4ews.util.Util.VerifyConnectionOnConnect"))) {
				logger.debug("Connection settings : trying to verify them");
				Folder.bind(service, WellKnownFolderName.Inbox);
				logger.info("Connection settings verified.");
			} else {
				logger.info("Connection settings not verified yet.");
			}
			return service;
		} catch (Exception e) {
		    Throwable cause = e.getCause();
		    if(cause != null) {
		        if (cause instanceof ConnectException) {
		            Exception nested = (ConnectException) cause;
		            throw new MessagingException(nested.getMessage(), nested);
                }
		    }
		    throw new AuthenticationFailedException(e.getMessage());
		}
	}

    private static Integer getConnectionTimeout(Session pSession) {
        Integer connectionTimeout = null;
        String cnxTimeoutStr = pSession.getProperty("mail.pop3.connectiontimeout");
        if(cnxTimeoutStr != null) {
            connectionTimeout = Integer.valueOf(cnxTimeoutStr);
        }
        return connectionTimeout;
    }

    private static Integer getProtocolTimeout(Session pSession) {
        Integer protocolTimeout = null;
        String protTimeoutStr = pSession.getProperty("mail.pop3.timeout");
        if( protTimeoutStr != null) {
            protocolTimeout = Integer.valueOf(protTimeoutStr);
        }
        return protocolTimeout;
    }
}
