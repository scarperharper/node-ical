/* eslint-disable max-depth, max-params, no-warning-comments, complexity */

const {v4: uuid} = require('uuid');
const moment = require('moment-timezone');
const rrule = require('rrule').RRule;

/** **************
 *  A tolerant, minimal icalendar parser
 *  (http://tools.ietf.org/html/rfc5545)
 *
 *  <peterbraden@peterbraden.co.uk>
 * ************* */

// Unescape Text re RFC 4.3.11
const text = function (t = '') {
  return t
    .replace(/\\,/g, ',')
    .replace(/\\;/g, ';')
    .replace(/\\[nN]/g, '\n')
    .replace(/\\\\/g, '\\');
};

const parseValue = function (value) {
  if (value === 'TRUE') {
    return true;
  }

  if (value === 'FALSE') {
    return false;
  }

  const number = Number(value);
  if (!Number.isNaN(number)) {
    return number;
  }

  return value;
};

const parseParameters = function (p) {
  const out = {};
  for (const element of p) {
    if (element.includes('=')) {
      const segs = element.split('=');

      out[segs[0]] = parseValue(segs.slice(1).join('='));
    }
  }

  // Sp is not defined in this scope, typo?
  // original code from peterbraden
  // return out || sp;
  return out;
};

const storeValueParameter = function (name) {
  return function (value, curr) {
    const current = curr[name];

    if (Array.isArray(current)) {
      current.push(value);
      return curr;
    }

    if (typeof current === 'undefined') {
      curr[name] = value;
    } else {
      curr[name] = [current, value];
    }

    return curr;
  };
};

const storeParameter = function (name) {
  return function (value, parameters, curr) {
    const data = parameters && parameters.length > 0 && !(parameters.length === 1 && parameters[0] === 'CHARSET=utf-8') ? {params: parseParameters(parameters), val: text(value)} : text(value);

    return storeValueParameter(name)(data, curr);
  };
};

const addTZ = function (dt, parameters) {
  const p = parseParameters(parameters);

  if (dt.tz) {
    // Date already has a timezone property
    return dt;
  }

  if (parameters && p && dt) {
    dt.tz = p.TZID;
    if (dt.tz !== undefined) {
      // Remove surrouding quotes if found at the begining and at the end of the string
      // (Occurs when parsing Microsoft Exchange events containing TZID with Windows standard format instead IANA)
      dt.tz = dt.tz.replace(/^"(.*)"$/, '$1');
    }
  }

  return dt;
};

let zoneTable = null;
function getIanaTZFromMS(msTZName) {
  if (!zoneTable) {
    zoneTable = JSON.parse(`{
		"Dateline Standard Time": {
			"iana": [
				"Etc/GMT+12"
			]
		},
		"UTC-11": {
			"iana": [
				"Etc/GMT+11"
			]
		},
		"Aleutian Standard Time": {
			"iana": [
				"America/Adak"
			]
		},
		"Hawaiian Standard Time": {
			"iana": [
				"Pacific/Honolulu"
			]
		},
		"Marquesas Standard Time": {
			"iana": [
				"Pacific/Marquesas"
			]
		},
		"Alaskan Standard Time": {
			"iana": [
				"America/Anchorage"
			]
		},
		"UTC-09": {
			"iana": [
				"Etc/GMT+9"
			]
		},
		"Pacific Standard Time (Mexico)": {
			"iana": [
				"America/Tijuana"
			]
		},
		"UTC-08": {
			"iana": [
				"Etc/GMT+8"
			]
		},
		"Pacific Standard Time": {
			"iana": [
				"America/Los_Angeles"
			]
		},
		"US Mountain Standard Time": {
			"iana": [
				"America/Phoenix"
			]
		},
		"Mountain Standard Time (Mexico)": {
			"iana": [
				"America/Chihuahua"
			]
		},
		"Mountain Standard Time": {
			"iana": [
				"America/Denver"
			]
		},
		"Yukon Standard Time": {
			"iana": [
				"America/Whitehorse"
			]
		},
		"Central America Standard Time": {
			"iana": [
				"America/Guatemala"
			]
		},
		"Central Standard Time": {
			"iana": [
				"America/Chicago"
			]
		},
		"Easter Island Standard Time": {
			"iana": [
				"Pacific/Easter"
			]
		},
		"Central Standard Time (Mexico)": {
			"iana": [
				"America/Mexico_City"
			]
		},
		"Canada Central Standard Time": {
			"iana": [
				"America/Regina"
			]
		},
		"SA Pacific Standard Time": {
			"iana": [
				"America/Bogota"
			]
		},
		"Eastern Standard Time (Mexico)": {
			"iana": [
				"America/Cancun"
			]
		},
		"Eastern Standard Time": {
			"iana": [
				"America/New_York"
			]
		},
		"Haiti Standard Time": {
			"iana": [
				"America/Port-au-Prince"
			]
		},
		"Cuba Standard Time": {
			"iana": [
				"America/Havana"
			]
		},
		"US Eastern Standard Time": {
			"iana": [
				"America/Indianapolis"
			]
		},
		"Turks And Caicos Standard Time": {
			"iana": [
				"America/Grand_Turk"
			]
		},
		"Paraguay Standard Time": {
			"iana": [
				"America/Asuncion"
			]
		},
		"Atlantic Standard Time": {
			"iana": [
				"America/Halifax"
			]
		},
		"Venezuela Standard Time": {
			"iana": [
				"America/Caracas"
			]
		},
		"Central Brazilian Standard Time": {
			"iana": [
				"America/Cuiaba"
			]
		},
		"SA Western Standard Time": {
			"iana": [
				"America/La_Paz"
			]
		},
		"Pacific SA Standard Time": {
			"iana": [
				"America/Santiago"
			]
		},
		"Newfoundland Standard Time": {
			"iana": [
				"America/St_Johns"
			]
		},
		"Tocantins Standard Time": {
			"iana": [
				"America/Araguaina"
			]
		},
		"E. South America Standard Time": {
			"iana": [
				"America/Sao_Paulo"
			]
		},
		"SA Eastern Standard Time": {
			"iana": [
				"America/Cayenne"
			]
		},
		"Argentina Standard Time": {
			"iana": [
				"America/Buenos_Aires"
			]
		},
		"Greenland Standard Time": {
			"iana": [
				"America/Godthab"
			]
		},
		"Montevideo Standard Time": {
			"iana": [
				"America/Montevideo"
			]
		},
		"Magallanes Standard Time": {
			"iana": [
				"America/Punta_Arenas"
			]
		},
		"Saint Pierre Standard Time": {
			"iana": [
				"America/Miquelon"
			]
		},
		"Bahia Standard Time": {
			"iana": [
				"America/Bahia"
			]
		},
		"UTC-02": {
			"iana": [
				"Etc/GMT+2"
			]
		},
		"Azores Standard Time": {
			"iana": [
				"Atlantic/Azores"
			]
		},
		"Cape Verde Standard Time": {
			"iana": [
				"Atlantic/Cape_Verde"
			]
		},
		"UTC": {
			"iana": [
				"Etc/UTC"
			]
		},
		"GMT Standard Time": {
			"iana": [
				"Europe/London"
			]
		},
		"Greenwich Standard Time": {
			"iana": [
				"Atlantic/Reykjavik"
			]
		},
		"Sao Tome Standard Time": {
			"iana": [
				"Africa/Sao_Tome"
			]
		},
		"Morocco Standard Time": {
			"iana": [
				"Africa/Casablanca"
			]
		},
		"W. Europe Standard Time": {
			"iana": [
				"Europe/Berlin"
			]
		},
		"Central Europe Standard Time": {
			"iana": [
				"Europe/Budapest"
			]
		},
		"Romance Standard Time": {
			"iana": [
				"Europe/Paris"
			]
		},
		"Central European Standard Time": {
			"iana": [
				"Europe/Warsaw"
			]
		},
		"W. Central Africa Standard Time": {
			"iana": [
				"Africa/Lagos"
			]
		},
		"Jordan Standard Time": {
			"iana": [
				"Asia/Amman"
			]
		},
		"GTB Standard Time": {
			"iana": [
				"Europe/Bucharest"
			]
		},
		"Middle East Standard Time": {
			"iana": [
				"Asia/Beirut"
			]
		},
		"Egypt Standard Time": {
			"iana": [
				"Africa/Cairo"
			]
		},
		"E. Europe Standard Time": {
			"iana": [
				"Europe/Chisinau"
			]
		},
		"Syria Standard Time": {
			"iana": [
				"Asia/Damascus"
			]
		},
		"West Bank Standard Time": {
			"iana": [
				"Asia/Hebron"
			]
		},
		"South Africa Standard Time": {
			"iana": [
				"Africa/Johannesburg"
			]
		},
		"FLE Standard Time": {
			"iana": [
				"Europe/Kiev"
			]
		},
		"Israel Standard Time": {
			"iana": [
				"Asia/Jerusalem"
			]
		},
		"South Sudan Standard Time": {
			"iana": [
				"Africa/Juba"
			]
		},
		"Kaliningrad Standard Time": {
			"iana": [
				"Europe/Kaliningrad"
			]
		},
		"Sudan Standard Time": {
			"iana": [
				"Africa/Khartoum"
			]
		},
		"Libya Standard Time": {
			"iana": [
				"Africa/Tripoli"
			]
		},
		"Namibia Standard Time": {
			"iana": [
				"Africa/Windhoek"
			]
		},
		"Arabic Standard Time": {
			"iana": [
				"Asia/Baghdad"
			]
		},
		"Turkey Standard Time": {
			"iana": [
				"Europe/Istanbul"
			]
		},
		"Arab Standard Time": {
			"iana": [
				"Asia/Riyadh"
			]
		},
		"Belarus Standard Time": {
			"iana": [
				"Europe/Minsk"
			]
		},
		"Russian Standard Time": {
			"iana": [
				"Europe/Moscow"
			]
		},
		"E. Africa Standard Time": {
			"iana": [
				"Africa/Nairobi"
			]
		},
		"Iran Standard Time": {
			"iana": [
				"Asia/Tehran"
			]
		},
		"Arabian Standard Time": {
			"iana": [
				"Asia/Dubai"
			]
		},
		"Astrakhan Standard Time": {
			"iana": [
				"Europe/Astrakhan"
			]
		},
		"Azerbaijan Standard Time": {
			"iana": [
				"Asia/Baku"
			]
		},
		"Russia Time Zone 3": {
			"iana": [
				"Europe/Samara"
			]
		},
		"Mauritius Standard Time": {
			"iana": [
				"Indian/Mauritius"
			]
		},
		"Saratov Standard Time": {
			"iana": [
				"Europe/Saratov"
			]
		},
		"Georgian Standard Time": {
			"iana": [
				"Asia/Tbilisi"
			]
		},
		"Volgograd Standard Time": {
			"iana": [
				"Europe/Volgograd"
			]
		},
		"Caucasus Standard Time": {
			"iana": [
				"Asia/Yerevan"
			]
		},
		"Afghanistan Standard Time": {
			"iana": [
				"Asia/Kabul"
			]
		},
		"West Asia Standard Time": {
			"iana": [
				"Asia/Tashkent"
			]
		},
		"Ekaterinburg Standard Time": {
			"iana": [
				"Asia/Yekaterinburg"
			]
		},
		"Pakistan Standard Time": {
			"iana": [
				"Asia/Karachi"
			]
		},
		"Qyzylorda Standard Time": {
			"iana": [
				"Asia/Qyzylorda"
			]
		},
		"India Standard Time": {
			"iana": [
				"Asia/Calcutta"
			]
		},
		"Sri Lanka Standard Time": {
			"iana": [
				"Asia/Colombo"
			]
		},
		"Nepal Standard Time": {
			"iana": [
				"Asia/Katmandu"
			]
		},
		"Central Asia Standard Time": {
			"iana": [
				"Asia/Almaty"
			]
		},
		"Bangladesh Standard Time": {
			"iana": [
				"Asia/Dhaka"
			]
		},
		"Omsk Standard Time": {
			"iana": [
				"Asia/Omsk"
			]
		},
		"Myanmar Standard Time": {
			"iana": [
				"Asia/Rangoon"
			]
		},
		"SE Asia Standard Time": {
			"iana": [
				"Asia/Bangkok"
			]
		},
		"Altai Standard Time": {
			"iana": [
				"Asia/Barnaul"
			]
		},
		"W. Mongolia Standard Time": {
			"iana": [
				"Asia/Hovd"
			]
		},
		"North Asia Standard Time": {
			"iana": [
				"Asia/Krasnoyarsk"
			]
		},
		"N. Central Asia Standard Time": {
			"iana": [
				"Asia/Novosibirsk"
			]
		},
		"Tomsk Standard Time": {
			"iana": [
				"Asia/Tomsk"
			]
		},
		"China Standard Time": {
			"iana": [
				"Asia/Shanghai"
			]
		},
		"North Asia East Standard Time": {
			"iana": [
				"Asia/Irkutsk"
			]
		},
		"Singapore Standard Time": {
			"iana": [
				"Asia/Singapore"
			]
		},
		"W. Australia Standard Time": {
			"iana": [
				"Australia/Perth"
			]
		},
		"Taipei Standard Time": {
			"iana": [
				"Asia/Taipei"
			]
		},
		"Ulaanbaatar Standard Time": {
			"iana": [
				"Asia/Ulaanbaatar"
			]
		},
		"Aus Central W. Standard Time": {
			"iana": [
				"Australia/Eucla"
			]
		},
		"Transbaikal Standard Time": {
			"iana": [
				"Asia/Chita"
			]
		},
		"Tokyo Standard Time": {
			"iana": [
				"Asia/Tokyo"
			]
		},
		"North Korea Standard Time": {
			"iana": [
				"Asia/Pyongyang"
			]
		},
		"Korea Standard Time": {
			"iana": [
				"Asia/Seoul"
			]
		},
		"Yakutsk Standard Time": {
			"iana": [
				"Asia/Yakutsk"
			]
		},
		"Cen. Australia Standard Time": {
			"iana": [
				"Australia/Adelaide"
			]
		},
		"AUS Central Standard Time": {
			"iana": [
				"Australia/Darwin"
			]
		},
		"E. Australia Standard Time": {
			"iana": [
				"Australia/Brisbane"
			]
		},
		"AUS Eastern Standard Time": {
			"iana": [
				"Australia/Sydney"
			]
		},
		"West Pacific Standard Time": {
			"iana": [
				"Pacific/Port_Moresby"
			]
		},
		"Tasmania Standard Time": {
			"iana": [
				"Australia/Hobart"
			]
		},
		"Vladivostok Standard Time": {
			"iana": [
				"Asia/Vladivostok"
			]
		},
		"Lord Howe Standard Time": {
			"iana": [
				"Australia/Lord_Howe"
			]
		},
		"Bougainville Standard Time": {
			"iana": [
				"Pacific/Bougainville"
			]
		},
		"Russia Time Zone 10": {
			"iana": [
				"Asia/Srednekolymsk"
			]
		},
		"Magadan Standard Time": {
			"iana": [
				"Asia/Magadan"
			]
		},
		"Norfolk Standard Time": {
			"iana": [
				"Pacific/Norfolk"
			]
		},
		"Sakhalin Standard Time": {
			"iana": [
				"Asia/Sakhalin"
			]
		},
		"Central Pacific Standard Time": {
			"iana": [
				"Pacific/Guadalcanal"
			]
		},
		"Russia Time Zone 11": {
			"iana": [
				"Asia/Kamchatka"
			]
		},
		"New Zealand Standard Time": {
			"iana": [
				"Pacific/Auckland"
			]
		},
		"UTC+12": {
			"iana": [
				"Etc/GMT-12"
			]
		},
		"Fiji Standard Time": {
			"iana": [
				"Pacific/Fiji"
			]
		},
		"Chatham Islands Standard Time": {
			"iana": [
				"Pacific/Chatham"
			]
		},
		"UTC+13": {
			"iana": [
				"Etc/GMT-13"
			]
		},
		"Tonga Standard Time": {
			"iana": [
				"Pacific/Tongatapu"
			]
		},
		"Samoa Standard Time": {
			"iana": [
				"Pacific/Apia"
			]
		},
		"Line Islands Standard Time": {
			"iana": [
				"Pacific/Kiritimati"
			]
		},
		"(UTC-12:00) International Date Line West": {
			"iana": [
				"Etc/GMT+12"
			]
		},
		"(UTC-11:00) Midway Island, Samoa": {
			"iana": [
				"Pacific/Apia"
			]
		},
		"(UTC-10:00) Hawaii": {
			"iana": [
				"Pacific/Honolulu"
			]
		},
		"(UTC-09:00) Alaska": {
			"iana": [
				"America/Anchorage"
			]
		},
		"(UTC-08:00) Pacific Time (US & Canada); Tijuana": {
			"iana": [
				"America/Los_Angeles"
			]
		},
		"(UTC-08:00) Pacific Time (US and Canada); Tijuana": {
			"iana": [
				"America/Los_Angeles"
			]
		},
		"(UTC-07:00) Mountain Time (US & Canada)": {
			"iana": [
				"America/Denver"
			]
		},
		"(UTC-07:00) Mountain Time (US and Canada)": {
			"iana": [
				"America/Denver"
			]
		},
		"(UTC-07:00) Chihuahua, La Paz, Mazatlan": {
			"iana": [
				null
			]
		},
		"(UTC-07:00) Arizona": {
			"iana": [
				"America/Phoenix"
			]
		},
		"(UTC-06:00) Central Time (US & Canada)": {
			"iana": [
				"America/Chicago"
			]
		},
		"(UTC-06:00) Central Time (US and Canada)": {
			"iana": [
				"America/Chicago"
			]
		},
		"(UTC-06:00) Saskatchewan": {
			"iana": [
				"America/Regina"
			]
		},
		"(UTC-06:00) Guadalajara, Mexico City, Monterrey": {
			"iana": [
				null
			]
		},
		"(UTC-06:00) Central America": {
			"iana": [
				"America/Guatemala"
			]
		},
		"(UTC-05:00) Eastern Time (US & Canada)": {
			"iana": [
				"America/New_York"
			]
		},
		"(UTC-05:00) Eastern Time (US and Canada)": {
			"iana": [
				"America/New_York"
			]
		},
		"(UTC-05:00) Indiana (East)": {
			"iana": [
				"America/Indianapolis"
			]
		},
		"(UTC-05:00) Bogota, Lima, Quito": {
			"iana": [
				"America/Bogota"
			]
		},
		"(UTC-04:00) Atlantic Time (Canada)": {
			"iana": [
				"America/Halifax"
			]
		},
		"(UTC-04:00) Georgetown, La Paz, San Juan": {
			"iana": [
				"America/La_Paz"
			]
		},
		"(UTC-04:00) Santiago": {
			"iana": [
				"America/Santiago"
			]
		},
		"(UTC-03:30) Newfoundland": {
			"iana": [
				null
			]
		},
		"(UTC-03:00) Brasilia": {
			"iana": [
				"America/Sao_Paulo"
			]
		},
		"(UTC-03:00) Georgetown": {
			"iana": [
				"America/Cayenne"
			]
		},
		"(UTC-03:00) Greenland": {
			"iana": [
				"America/Godthab"
			]
		},
		"(UTC-02:00) Mid-Atlantic": {
			"iana": [
				null
			]
		},
		"(UTC-01:00) Azores": {
			"iana": [
				"Atlantic/Azores"
			]
		},
		"(UTC-01:00) Cape Verde Islands": {
			"iana": [
				"Atlantic/Cape_Verde"
			]
		},
		"(UTC) Greenwich Mean Time: Dublin, Edinburgh, Lisbon, London": {
			"iana": [
				null
			]
		},
		"(UTC) Monrovia, Reykjavik": {
			"iana": [
				"Atlantic/Reykjavik"
			]
		},
		"(UTC+01:00) Belgrade, Bratislava, Budapest, Ljubljana, Prague": {
			"iana": [
				"Europe/Budapest"
			]
		},
		"(UTC+01:00) Sarajevo, Skopje, Warsaw, Zagreb": {
			"iana": [
				"Europe/Warsaw"
			]
		},
		"(UTC+01:00) Brussels, Copenhagen, Madrid, Paris": {
			"iana": [
				"Europe/Paris"
			]
		},
		"(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna": {
			"iana": [
				"Europe/Berlin"
			]
		},
		"(UTC+01:00) West Central Africa": {
			"iana": [
				"Africa/Lagos"
			]
		},
		"(UTC+02:00) Minsk": {
			"iana": [
				"Europe/Chisinau"
			]
		},
		"(UTC+02:00) Cairo": {
			"iana": [
				"Africa/Cairo"
			]
		},
		"(UTC+02:00) Helsinki, Kiev, Riga, Sofia, Tallinn, Vilnius": {
			"iana": [
				"Europe/Kiev"
			]
		},
		"(UTC+02:00) Athens, Bucharest, Istanbul": {
			"iana": [
				"Europe/Bucharest"
			]
		},
		"(UTC+02:00) Jerusalem": {
			"iana": [
				"Asia/Jerusalem"
			]
		},
		"(UTC+02:00) Harare, Pretoria": {
			"iana": [
				"Africa/Johannesburg"
			]
		},
		"(UTC+03:00) Moscow, St. Petersburg, Volgograd": {
			"iana": [
				"Europe/Moscow"
			]
		},
		"(UTC+03:00) Kuwait, Riyadh": {
			"iana": [
				"Asia/Riyadh"
			]
		},
		"(UTC+03:00) Nairobi": {
			"iana": [
				"Africa/Nairobi"
			]
		},
		"(UTC+03:00) Baghdad": {
			"iana": [
				"Asia/Baghdad"
			]
		},
		"(UTC+03:30) Tehran": {
			"iana": [
				"Asia/Tehran"
			]
		},
		"(UTC+04:00) Abu Dhabi, Muscat": {
			"iana": [
				"Asia/Dubai"
			]
		},
		"(UTC+04:00) Baku, Tbilisi, Yerevan": {
			"iana": [
				"Asia/Yerevan"
			]
		},
		"(UTC+04:30) Kabul": {
			"iana": [
				null
			]
		},
		"(UTC+05:00) Ekaterinburg": {
			"iana": [
				"Asia/Yekaterinburg"
			]
		},
		"(UTC+05:00) Tashkent": {
			"iana": [
				"Asia/Tashkent"
			]
		},
		"(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi": {
			"iana": [
				"Asia/Calcutta"
			]
		},
		"(UTC+05:45) Kathmandu": {
			"iana": [
				"Asia/Katmandu"
			]
		},
		"(UTC+06:00) Astana, Dhaka": {
			"iana": [
				"Asia/Almaty"
			]
		},
		"(UTC+06:00) Sri Jayawardenepura": {
			"iana": [
				"Asia/Colombo"
			]
		},
		"(UTC+06:00) Almaty, Novosibirsk": {
			"iana": [
				"Asia/Novosibirsk"
			]
		},
		"(UTC+06:30) Yangon (Rangoon)": {
			"iana": [
				"Asia/Rangoon"
			]
		},
		"(UTC+07:00) Bangkok, Hanoi, Jakarta": {
			"iana": [
				"Asia/Bangkok"
			]
		},
		"(UTC+07:00) Krasnoyarsk": {
			"iana": [
				"Asia/Krasnoyarsk"
			]
		},
		"(UTC+08:00) Beijing, Chongqing, Hong Kong, Urumqi": {
			"iana": [
				"Asia/Shanghai"
			]
		},
		"(UTC+08:00) Kuala Lumpur, Singapore": {
			"iana": [
				"Asia/Singapore"
			]
		},
		"(UTC+08:00) Taipei": {
			"iana": [
				"Asia/Taipei"
			]
		},
		"(UTC+08:00) Perth": {
			"iana": [
				"Australia/Perth"
			]
		},
		"(UTC+08:00) Irkutsk, Ulaanbaatar": {
			"iana": [
				"Asia/Irkutsk"
			]
		},
		"(UTC+09:00) Seoul": {
			"iana": [
				"Asia/Seoul"
			]
		},
		"(UTC+09:00) Osaka, Sapporo, Tokyo": {
			"iana": [
				"Asia/Tokyo"
			]
		},
		"(UTC+09:00) Yakutsk": {
			"iana": [
				"Asia/Yakutsk"
			]
		},
		"(UTC+09:30) Darwin": {
			"iana": [
				"Australia/Darwin"
			]
		},
		"(UTC+09:30) Adelaide": {
			"iana": [
				"Australia/Adelaide"
			]
		},
		"(UTC+10:00) Canberra, Melbourne, Sydney": {
			"iana": [
				"Australia/Sydney"
			]
		},
		"(GMT+10:00) Canberra, Melbourne, Sydney": {
			"iana": [
				"Australia/Sydney"
			]
		},
		"(UTC+10:00) Brisbane": {
			"iana": [
				"Australia/Brisbane"
			]
		},
		"(UTC+10:00) Hobart": {
			"iana": [
				"Australia/Hobart"
			]
		},
		"(UTC+10:00) Vladivostok": {
			"iana": [
				"Asia/Vladivostok"
			]
		},
		"(UTC+10:00) Guam, Port Moresby": {
			"iana": [
				"Pacific/Port_Moresby"
			]
		},
		"(UTC+11:00) Magadan, Solomon Islands, New Caledonia": {
			"iana": [
				"Pacific/Guadalcanal"
			]
		},
		"(UTC+12:00) Fiji, Kamchatka, Marshall Is.": {
			"iana": [
				null
			]
		},
		"(UTC+12:00) Auckland, Wellington": {
			"iana": [
				"Pacific/Auckland"
			]
		},
		"(UTC+13:00) Nukualofa": {
			"iana": [
				"Pacific/Tongatapu"
			]
		},
		"(UTC-03:00) Buenos Aires": {
			"iana": [
				"America/Buenos_Aires"
			]
		},
		"(UTC+02:00) Beirut": {
			"iana": [
				"Asia/Beirut"
			]
		},
		"(UTC+02:00) Amman": {
			"iana": [
				"Asia/Amman"
			]
		},
		"(UTC-06:00) Guadalajara, Mexico City, Monterrey - New": {
			"iana": [
				"America/Mexico_City"
			]
		},
		"(UTC-07:00) Chihuahua, La Paz, Mazatlan - New": {
			"iana": [
				"America/Chihuahua"
			]
		},
		"(UTC-08:00) Tijuana, Baja California": {
			"iana": [
				"America/Tijuana"
			]
		},
		"(UTC+02:00) Windhoek": {
			"iana": [
				"Africa/Windhoek"
			]
		},
		"(UTC+03:00) Tbilisi": {
			"iana": [
				"Asia/Tbilisi"
			]
		},
		"(UTC-04:00) Manaus": {
			"iana": [
				"America/Cuiaba"
			]
		},
		"(UTC-03:00) Montevideo": {
			"iana": [
				"America/Montevideo"
			]
		},
		"(UTC+04:00) Yerevan": {
			"iana": [
				null
			]
		},
		"(UTC-04:30) Caracas": {
			"iana": [
				"America/Caracas"
			]
		},
		"(UTC) Casablanca": {
			"iana": [
				"Africa/Casablanca"
			]
		},
		"(UTC+05:00) Islamabad, Karachi": {
			"iana": [
				"Asia/Karachi"
			]
		},
		"(UTC+04:00) Port Louis": {
			"iana": [
				"Indian/Mauritius"
			]
		},
		"(UTC) Coordinated Universal Time": {
			"iana": [
				"Etc/UTC"
			]
		},
		"(UTC-04:00) Asuncion": {
			"iana": [
				"America/Asuncion"
			]
		},
		"(UTC+12:00) Petropavlovsk-Kamchatsky": {
			"iana": [
				null
			]
		}
	}`)
  }

  // Get hash entry
  const he = zoneTable[msTZName];
  // If found return iana name, else null
  return he ? he.iana[0] : null;
}

function getTimeZone(value) {
  let tz = value;
  let found = '';
  // If this is the custom timezone from MS Outlook
  if (tz === 'tzone://Microsoft/Custom') {
    // Set it to the local timezone, cause we can't tell
    tz = moment.tz.guess();
  }

  // Remove quotes if found
  tz = tz.replace(/^"(.*)"$/, '$1');

  // Watch out for windows timezones
  if (tz && tz.includes(' ')) {
    const tz1 = getIanaTZFromMS(tz);
    if (tz1) {
      tz = tz1;
    }
  }

  // Watch out for offset timezones
  // If the conversion above didn't find any matching IANA tz
  // And offset is still present
  if (tz && tz.startsWith('(')) {
    // Extract just the offset
    const regex = /[+|-]\d*:\d*/;
    tz = null;
    found = tz.match(regex);
  }

  // Timezone not confirmed yet
  if (found === '') {
    // Lookup tz
    found = moment.tz.names().find(zone => {
      return zone === tz;
    });
  }

  return found === '' ? tz : found;
}

function isDateOnly(value, parameters) {
  const dateOnly = ((parameters && parameters.includes('VALUE=DATE') && !parameters.includes('VALUE=DATE-TIME')) || /^\d{8}$/.test(value) === true);
  return dateOnly;
}

const typeParameter = function (name) {
  // Typename is not used in this function?
  return function (value, parameters, curr) {
    const returnValue = isDateOnly(value, parameters) ? 'date' : 'date-time';
    return storeValueParameter(name)(returnValue, curr);
  };
};

const dateParameter = function (name) {
  return function (value, parameters, curr) {
    // The regex from main gets confued by extra :
    const pi = parameters.indexOf('TZID=tzone');
    if (pi >= 0) {
      // Correct the parameters with the part on the value
      parameters[pi] = parameters[pi] + ':' + value.split(':')[0];
      // Get the date from the field, other code uses the value parameter
      value = value.split(':')[1];
    }

    let newDate = text(value);

    // Process 'VALUE=DATE' and EXDATE
    if (isDateOnly(value, parameters)) {
      // Just Date

      const comps = /^(\d{4})(\d{2})(\d{2}).*$/.exec(value);
      if (comps !== null) {
        // No TZ info - assume same timezone as this computer
        newDate = new Date(comps[1], Number.parseInt(comps[2], 10) - 1, comps[3]);

        newDate.dateOnly = true;

        // Store as string - worst case scenario
        return storeValueParameter(name)(newDate, curr);
      }
    }

    // Typical RFC date-time format
    const comps = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z)?$/.exec(value);
    if (comps !== null) {
      if (comps[7] === 'Z') {
        // GMT
        newDate = new Date(
          Date.UTC(
            Number.parseInt(comps[1], 10),
            Number.parseInt(comps[2], 10) - 1,
            Number.parseInt(comps[3], 10),
            Number.parseInt(comps[4], 10),
            Number.parseInt(comps[5], 10),
            Number.parseInt(comps[6], 10)
          )
        );
        newDate.tz = 'Etc/UTC';
      } else if (parameters && parameters[0] && parameters[0].includes('TZID=') && parameters[0].split('=')[1]) {
        // Get the timeozone from trhe parameters TZID value
        let tz = parameters[0].split('=')[1];
        let found = '';
        let offset = '';

        // If this is the custom timezone from MS Outlook
        if (tz === 'tzone://Microsoft/Custom') {
          // Set it to the local timezone, cause we can't tell
          tz = moment.tz.guess();
          parameters[0] = 'TZID=' + tz;
        }

        // Remove quotes if found
        tz = tz.replace(/^"(.*)"$/, '$1');

        // Watch out for windows timezones
        if (tz && tz.includes(' ')) {
          const tz1 = getIanaTZFromMS(tz);
          if (tz1) {
            tz = tz1;
            // We have a confirmed timezone, dont use offset, may confuse DST/STD time
            offset = '';
          }
        }

        // Watch out for offset timezones
        // If the conversion above didn't find any matching IANA tz
        // And oiffset is still present
        if (tz && tz.startsWith('(')) {
          // Extract just the offset
          const regex = /[+|-]\d*:\d*/;
          offset = tz.match(regex);
          tz = null;
          found = offset;
        }

        // Timezone not confirmed yet
        if (found === '') {
          // Lookup tz
          found = moment.tz.names().find(zone => {
            return zone === tz;
          });
        }

        // Timezone confirmed or forced to offset
        newDate = found ? moment.tz(value, 'YYYYMMDDTHHmmss' + offset, tz).toDate() : new Date(
          Number.parseInt(comps[1], 10),
          Number.parseInt(comps[2], 10) - 1,
          Number.parseInt(comps[3], 10),
          Number.parseInt(comps[4], 10),
          Number.parseInt(comps[5], 10),
          Number.parseInt(comps[6], 10)
        );

        newDate = addTZ(newDate, parameters);
      } else {
        newDate = new Date(
          Number.parseInt(comps[1], 10),
          Number.parseInt(comps[2], 10) - 1,
          Number.parseInt(comps[3], 10),
          Number.parseInt(comps[4], 10),
          Number.parseInt(comps[5], 10),
          Number.parseInt(comps[6], 10)
        );
      }
    }

    // Store as string - worst case scenario
    return storeValueParameter(name)(newDate, curr);
  };
};

const geoParameter = function (name) {
  return function (value, parameters, curr) {
    storeParameter(value, parameters, curr);
    const parts = value.split(';');
    curr[name] = {lat: Number(parts[0]), lon: Number(parts[1])};
    return curr;
  };
};

const categoriesParameter = function (name) {
  const separatorPattern = /\s*,\s*/g;
  return function (value, parameters, curr) {
    storeParameter(value, parameters, curr);
    if (curr[name] === undefined) {
      curr[name] = value ? value.split(separatorPattern) : [];
    } else if (value) {
      curr[name] = curr[name].concat(value.split(separatorPattern));
    }

    return curr;
  };
};

// EXDATE is an entry that represents exceptions to a recurrence rule (ex: "repeat every day except on 7/4").
// The EXDATE entry itself can also contain a comma-separated list, so we make sure to parse each date out separately.
// There can also be more than one EXDATE entries in a calendar record.
// Since there can be multiple dates, we create an array of them.  The index into the array is the ISO string of the date itself, for ease of use.
// i.e. You can check if ((curr.exdate != undefined) && (curr.exdate[date iso string] != undefined)) to see if a date is an exception.
// NOTE: This specifically uses date only, and not time.  This is to avoid a few problems:
//    1. The ISO string with time wouldn't work for "floating dates" (dates without timezones).
//       ex: "20171225T060000" - this is supposed to mean 6 AM in whatever timezone you're currently in
//    2. Daylight savings time potentially affects the time you would need to look up
//    3. Some EXDATE entries in the wild seem to have times different from the recurrence rule, but are still excluded by calendar programs.  Not sure how or why.
//       These would fail any sort of sane time lookup, because the time literally doesn't match the event.  So we'll ignore time and just use date.
//       ex: DTSTART:20170814T140000Z
//             RRULE:FREQ=WEEKLY;WKST=SU;INTERVAL=2;BYDAY=MO,TU
//             EXDATE:20171219T060000
//       Even though "T060000" doesn't match or overlap "T1400000Z", it's still supposed to be excluded?  Odd. :(
// TODO: See if this causes any problems with events that recur multiple times a day.
const exdateParameter = function (name) {
  return function (value, parameters, curr) {
    const separatorPattern = /\s*,\s*/g;
    curr[name] = curr[name] || [];
    const dates = value ? value.split(separatorPattern) : [];
    for (const entry of dates) {
      const exdate = [];
      dateParameter(name)(entry, parameters, exdate);

      if (exdate[name]) {
        if (typeof exdate[name].toISOString === 'function') {
          curr[name][exdate[name].toISOString().slice(0, 10)] = exdate[name];
        } else {
          throw new TypeError('No toISOString function in exdate[name]', exdate[name]);
        }
      }
    }

    return curr;
  };
};

// RECURRENCE-ID is the ID of a specific recurrence within a recurrence rule.
// TODO:  It's also possible for it to have a range, like "THISANDPRIOR", "THISANDFUTURE".  This isn't currently handled.
const recurrenceParameter = function (name) {
  return dateParameter(name);
};

const addFBType = function (fb, parameters) {
  const p = parseParameters(parameters);

  if (parameters && p) {
    fb.type = p.FBTYPE || 'BUSY';
  }

  return fb;
};

const freebusyParameter = function (name) {
  return function (value, parameters, curr) {
    const fb = addFBType({}, parameters);
    curr[name] = curr[name] || [];
    curr[name].push(fb);

    storeParameter(value, parameters, fb);

    const parts = value.split('/');

    for (const [index, name] of ['start', 'end'].entries()) {
      dateParameter(name)(parts[index], parameters, fb);
    }

    return curr;
  };
};

module.exports = {
  objectHandlers: {
    BEGIN(component, parameters, curr, stack) {
      stack.push(curr);

      return {type: component, params: parameters};
    },
    END(value, parameters, curr, stack) {
      // Original end function
      const originalEnd = function (component, parameters_, curr, stack) {
        // Prevents the need to search the root of the tree for the VCALENDAR object
        if (component === 'VCALENDAR') {
          // Scan all high level object in curr and drop all strings
          let key;
          let object;

          for (key in curr) {
            if (!{}.hasOwnProperty.call(curr, key)) {
              continue;
            }

            object = curr[key];
            if (typeof object === 'string') {
              delete curr[key];
            }
          }

          return curr;
        }

        const par = stack.pop();

        if (!curr.end) { // RFC5545, 3.6.1
          if (curr.datetype === 'date-time') {
            curr.end = curr.start;
            // If the duration is not set
          } else if (curr.duration === undefined) {
            // Set the end to the start plus one day RFC5545, 3.6.1
            curr.end = moment.utc(curr.start).add(1, 'days').toDate(); // New Date(moment(curr.start).add(1, 'days'));
          } else {
            const durationUnits =
              {
                // Y: 'years',
                // M: 'months',
                W: 'weeks',
                D: 'days',
                H: 'hours',
                M: 'minutes',
                S: 'seconds'
              };
            // Get the list of duration elements
            const r = curr.duration.match(/-?\d+[YMWDHS]/g);
            let newend = moment.utc(curr.start);
            // Is the 1st character a negative sign?
            const indicator = curr.duration.startsWith('-') ? -1 : 1;
            // Process each element
            for (const d of r) {
              newend = newend.add(Number.parseInt(d, 10) * indicator, durationUnits[d.slice(-1)]);
            }

            curr.end = newend.toDate();
          }
        }

        if (curr.uid) {
          // If this is the first time we run into this UID, just save it.
          if (par[curr.uid] === undefined) {
            par[curr.uid] = curr;

            if (par.method) { // RFC5545, 3.2
              par[curr.uid].method = par.method;
            }
          } else if (curr.recurrenceid === undefined) {
            // If we have multiple ical entries with the same UID, it's either going to be a
            // modification to a recurrence (RECURRENCE-ID), and/or a significant modification
            // to the entry (SEQUENCE).

            // TODO: Look into proper sequence logic.

            // If we have the same UID as an existing record, and it *isn't* a specific recurrence ID,
            // not quite sure what the correct behaviour should be.  For now, just take the new information
            // and merge it with the old record by overwriting only the fields that appear in the new record.
            let key;
            for (key in curr) {
              if (key !== null) {
                par[curr.uid][key] = curr[key];
              }
            }
          }

          // If we have recurrence-id entries, list them as an array of recurrences keyed off of recurrence-id.
          // To use - as you're running through the dates of an rrule, you can try looking it up in the recurrences
          // array.  If it exists, then use the data from the calendar object in the recurrence instead of the parent
          // for that day.

          // NOTE:  Sometimes the RECURRENCE-ID record will show up *before* the record with the RRULE entry.  In that
          // case, what happens is that the RECURRENCE-ID record ends up becoming both the parent record and an entry
          // in the recurrences array, and then when we process the RRULE entry later it overwrites the appropriate
          // fields in the parent record.

          if (typeof curr.recurrenceid !== 'undefined') {
            // TODO:  Is there ever a case where we have to worry about overwriting an existing entry here?

            // Create a copy of the current object to save in our recurrences array.  (We *could* just do par = curr,
            // except for the case that we get the RECURRENCE-ID record before the RRULE record.  In that case, we
            // would end up with a shared reference that would cause us to overwrite *both* records at the point
            // that we try and fix up the parent record.)
            const recurrenceObject = {};
            let key;
            for (key in curr) {
              if (key !== null) {
                recurrenceObject[key] = curr[key];
              }
            }

            if (typeof recurrenceObject.recurrences !== 'undefined') {
              delete recurrenceObject.recurrences;
            }

            // If we don't have an array to store recurrences in yet, create it.
            if (par[curr.uid].recurrences === undefined) {
              par[curr.uid].recurrences = {};
            }

            // Save off our cloned recurrence object into the array, keyed by date but not time.
            // We key by date only to avoid timezone and "floating time" problems (where the time isn't associated with a timezone).
            // TODO: See if this causes a problem with events that have multiple recurrences per day.
            if (typeof curr.recurrenceid.toISOString === 'function') {
              par[curr.uid].recurrences[curr.recurrenceid.toISOString().slice(0, 10)] = recurrenceObject;
            } else { // Removed issue 56
              throw new TypeError('No toISOString function in curr.recurrenceid', curr.recurrenceid);
            }
          }

          // One more specific fix - in the case that an RRULE entry shows up after a RECURRENCE-ID entry,
          // let's make sure to clear the recurrenceid off the parent field.
          if (typeof par[curr.uid].rrule !== 'undefined' && typeof par[curr.uid].recurrenceid !== 'undefined') {
            delete par[curr.uid].recurrenceid;
          }
        } else {
          const id = uuid();
          par[id] = curr;

          if (par.method) { // RFC5545, 3.2
            par[id].method = par.method;
          }
        }

        return par;
      };

      // Recurrence rules are only valid for VEVENT, VTODO, and VJOURNAL.
      // More specifically, we need to filter the VCALENDAR type because we might end up with a defined rrule
      // due to the subtypes.

      if ((value === 'VEVENT' || value === 'VTODO' || value === 'VJOURNAL') && curr.rrule) {
        let rule = curr.rrule.replace('RRULE:', '');
        // Make sure the rrule starts with FREQ=
        rule = rule.slice(rule.lastIndexOf('FREQ='));
        // If no rule start date
        if (rule.includes('DTSTART') === false) {
          // Get date/time into a specific format for comapare
          let x = moment(curr.start).format('MMMM/Do/YYYY, h:mm:ss a');
          // If the local time value is midnight
          // This a whole day event
          if (x.slice(-11) === '12:00:00 am') {
            // Get the timezone offset
            // The internal date is stored in UTC format
            const offset = curr.start.getTimezoneOffset();
            // Only east of gmt is a problem
            if (offset < 0) {
              // Calculate the new startdate with the offset applied, bypass RRULE/Luxon confusion
              // Make the internally stored DATE the actual date (not UTC offseted)
              // Luxon expects local time, not utc, so gets start date wrong if not adjusted
              curr.start = new Date(curr.start.getTime() + (Math.abs(offset) * 60000));
            } else {
              // Get rid of any time (shouldn't be any, but be sure)
              x = moment(curr.start).format('MMMM/Do/YYYY');
              const comps = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(x);
              if (comps) {
                curr.start = new Date(comps[3], comps[1] - 1, comps[2]);
              }
            }
          }

          // If the date has an toISOString function
          if (curr.start && typeof curr.start.toISOString === 'function') {
            try {
              // If the original date has a TZID, add it
              if (curr.start.tz) {
                const tz = getTimeZone(curr.start.tz);
                rule += `;DTSTART;TZID=${tz}:${curr.start.toISOString().replace(/[-:]/g, '')}`;
              } else {
                rule += `;DTSTART=${curr.start.toISOString().replace(/[-:]/g, '')}`;
              }

              rule = rule.replace(/\.\d{3}/, '');
            } catch (error) { // This should not happen, issue #56
              throw new Error('ERROR when trying to convert to ISOString', error);
            }
          } else {
            throw new Error('No toISOString function in curr.start', curr.start);
          }
        }

        // Make sure to catch error from rrule.fromString()
        try {
          curr.rrule = rrule.fromString(rule);
        } catch (error) {
          throw error;
        }
      }

      return originalEnd.call(this, value, parameters, curr, stack);
    },
    SUMMARY: storeParameter('summary'),
    DESCRIPTION: storeParameter('description'),
    URL: storeParameter('url'),
    UID: storeParameter('uid'),
    LOCATION: storeParameter('location'),
    DTSTART(value, parameters, curr) {
      curr = dateParameter('start')(value, parameters, curr);
      return typeParameter('datetype')(value, parameters, curr);
    },
    DTEND: dateParameter('end'),
    EXDATE: exdateParameter('exdate'),
    ' CLASS': storeParameter('class'), // Should there be a space in this property?
    TRANSP: storeParameter('transparency'),
    GEO: geoParameter('geo'),
    'PERCENT-COMPLETE': storeParameter('completion'),
    COMPLETED: dateParameter('completed'),
    CATEGORIES: categoriesParameter('categories'),
    FREEBUSY: freebusyParameter('freebusy'),
    DTSTAMP: dateParameter('dtstamp'),
    CREATED: dateParameter('created'),
    'LAST-MODIFIED': dateParameter('lastmodified'),
    'RECURRENCE-ID': recurrenceParameter('recurrenceid'),
    RRULE(value, parameters, curr, stack, line) {
      curr.rrule = line;
      return curr;
    }
  },

  handleObject(name, value, parameters, ctx, stack, line) {
    if (this.objectHandlers[name]) {
      return this.objectHandlers[name](value, parameters, ctx, stack, line);
    }

    // Handling custom properties
    if (/X-[\w-]+/.test(name) && stack.length > 0) {
      // Trimming the leading and perform storeParam
      name = name.slice(2);
      return storeParameter(name)(value, parameters, ctx, stack, line);
    }

    return storeParameter(name.toLowerCase())(value, parameters, ctx);
  },

  parseLines(lines, limit, ctx, stack, lastIndex, cb) {
    if (!cb && typeof ctx === 'function') {
      cb = ctx;
      ctx = undefined;
    }

    ctx = ctx || {};
    stack = stack || [];

    let limitCounter = 0;

    let i = lastIndex || 0;
    for (let ii = lines.length; i < ii; i++) {
      let l = lines[i];
      // Unfold : RFC#3.1
      while (lines[i + 1] && /[ \t]/.test(lines[i + 1][0])) {
        l += lines[i + 1].slice(1);
        i++;
      }

      // Remove any double quotes in any tzid statement// except around (utc+hh:mm
      if (l.indexOf('TZID=') && !l.includes('"(')) {
        l = l.replace(/"/g, '');
      }

      const exp = /^([\w\d-]+)((?:;[\w\d-]+=(?:(?:"[^"]*")|[^":;]+))*):(.*)$/;
      let kv = l.match(exp);

      if (kv === null) {
        // Invalid line - must have k&v
        continue;
      }

      kv = kv.slice(1);

      const value = kv[kv.length - 1];
      const name = kv[0];
      const parameters = kv[1] ? kv[1].split(';').slice(1) : [];

      ctx = this.handleObject(name, value, parameters, ctx, stack, l) || {};
      if (++limitCounter > limit) {
        break;
      }
    }

    if (i >= lines.length) {
      // Type and params are added to the list of items, get rid of them.
      delete ctx.type;
      delete ctx.params;
    }

    if (cb) {
      if (i < lines.length) {
        setImmediate(() => {
          this.parseLines(lines, limit, ctx, stack, i + 1, cb);
        });
      } else {
        setImmediate(() => {
          cb(null, ctx);
        });
      }
    } else {
      return ctx;
    }
  },

  getLineBreakChar(string) {
    const indexOfLF = string.indexOf('\n', 1); // No need to check first-character

    if (indexOfLF === -1) {
      if (string.includes('\r')) {
        return '\r';
      }

      return '\n';
    }

    if (string[indexOfLF - 1] === '\r') {
      return '\r?\n';
    }

    return '\n';
  },

  parseICS(string, cb) {
    const lineEndType = this.getLineBreakChar(string);
    const lines = string.split(lineEndType === '\n' ? /\n/ : /\r?\n/);
    let ctx;

    if (cb) {
      // Asynchronous execution
      this.parseLines(lines, 2000, cb);
    } else {
      // Synchronous execution
      ctx = this.parseLines(lines, lines.length);
      return ctx;
    }
  }
};
