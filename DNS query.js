/**
 * A quick port of these Apps Script custom functions >> https://github.com/pfelipm/fxdnsquery
 * 01/08/22 @pfelipm
 */

// Init
import * as coda from "@codahq/packs-sdk";
export const pack = coda.newPack();

// Allow external connections to: 
pack.addNetworkDomain("cloudflare-dns.com");

// Some global constants
const DNS_RECORDS = ["A", "AAAA", "CAA", "CNAME", "DS", "DNSKEY", "MX", "NS", "NSEC", "NSEC3", "RRSIG", "SOA", "TXT"];
const CLOUDFLARE_DNS_ENDPOINT = "https://cloudflare-dns.com/dns-query";

// Helper function, based on original code by Cloudflare:
// https://developers.cloudflare.com/1.1.1.1/other-ways-to-use-1.1.1.1/dns-in-google-sheets/
async function NSLookup(type, domain, context) {

  const errors = [
    { "name": "NoError", "description": "No Error." }, // 0
    { "name": "FormErr", "description": "Format Error." }, // 1
    { "name": "ServFail", "description": "Server Failure." }, // 2
    { "name": "NXDomain", "description": "Non-Existent Domain." }, // 3
    { "name": "NotImp", "description": "Not Implemented." }, // 4
    { "name": "Refused", "description": "Query Refused." }, // 5
    { "name": "YXDomain", "description": "Name Exists when it should not." }, // 6
    { "name": "YXRRSet", "description": "RR Set Exists when it should not." }, // 7
    { "name": "NXRRSet", "description": "RR Set that should exist does not." }, // 8
    { "name": "NotAuth", "description": "Not Authorized." } // 9
  ];

  try {

    const response = await context.fetcher.fetch({
      method: "GET",
      // cacheTtlSecs: 8 * 60 * 60;
      url: `${CLOUDFLARE_DNS_ENDPOINT}?name=${encodeURIComponent(domain)}&type=${encodeURIComponent(type)}`,
      headers: { accept: "application/dns-json" }
    });

    // Fetch OK?
    if (response.status != 200) throw new Error(`Internal fetch error: ${response.status}`);

    // Checks status of response object, quits if error reported (see table above)
    if (response.body.Status !== 0) return `Error: ${errors[response.body.Status].description}`;

    // Builds & returns non-error result
    const outputData = [];
    response.body.Answer.forEach(answer => outputData.push(answer.data));
    return outputData.join();

  } catch (e) {
    return "Could not fetch result!";
  }

}


// Formula that fetches a DNS record from the provided domain
pack.addFormula({

  name: "DnsRecord",
  description: "Looks up a DNS record of the provided domain",

  // Parameters
  parameters: [
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "dnsRecordType",
      description: "DNS record to fetch",
      autocomplete: DNS_RECORDS
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "domain",
      description: "Domain to query"
    }),
  ],

  // Result
  resultType: coda.ValueType.String,

  // Examples
  examples: [
    {
      params: ["MX", "google.es"],
      result: "0 smtp.google.com"
    },
    {
      params: ["SOA", "outlook.com"],
      result: "ch0mgt0101dc001.prdmgt01.prod.exchangelabs.com. msnhst.microsoft.com. 2015311244 300 900 2419200 60"
    },
    {
      params: ["A", "uji.es"],
      result: "150.128.98.231,150.128.98.232"
    }
  ],

  // Actual code
  execute: function ([dnsRecord, domain], context) {

    // No need to check param type, Coda takes care of converting it to the declared type, it seems
    dnsRecord = dnsRecord.toUpperCase().trim();
    domain = domain.toLowerCase().trim();

    // These messages will appear in the Packs execution log (click on "View error details"),
    // does not seem appropriate to show meaningful error messages to the user.
    if (dnsRecord.length == 0) throw new Error("Invalid DNS record type.");
    if (domain.length == 0) throw new Error("Invalid domain.");
    if (!DNS_RECORDS.includes(dnsRecord)) throw new Error("Unknown dnsRecordType.");

    return NSLookup(dnsRecord, domain, context);

  }

});

// IsGoogleEmail formula below will be used in a column format
pack.addColumnFormat({

  name: "Is a Google email",
  instructions: "Will return a true value if email addresses (or domains) in this column are of the Gmail or Google Workspace type.",
  formulaName: "IsGoogleEmail"

});

// Formula that checks whether an email/domain belongs to a consumer or Workspace Google account
pack.addFormula({

  name: "IsGoogleEmail",
  description: "Finds out if an email address (or domain) is of the Gmail or Google Workspace type",

  // Parameters
  parameters: [
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "emailAddress",
      description: "Email address or domain to test",
    }),
    coda.makeParameter({
      type: coda.ParameterType.String,
      name: "type",
      description: "Type of check, if not specified or unknown tests both personal and Workspace email types",
      autocomplete: ["gmail", "workspace", 'google'],
      optional: true
    }),
  ],

  // Result
  resultType: coda.ValueType.Boolean,

  // Examples
  examples: [
    {
      params: ["takerna@gmail.com"],
      result: "true"
    },
    {
      params: ["coordinacion@gedu.es", "workspace"],
      result: "true"
    },
    {
      params: ["pablo@masmenos.tk", "google"],
      result: "false"
    }
  ],

  // Actual code
  execute: function ([email, testType = "google"], context) {

    email = email.toLowerCase().trim();
    testType = testType.toLowerCase().trim();

    // No need to check param type, Coda takes care of converting it to the declared type, it seems
    if (email.length == 0) throw new Error("Invalid DNS record type.");

    let domains = [];
    switch (testType) {
      case "workspace":
        domains = ["aspmx.l.google.com", "googlemail.com"]; // 2nd one is obsolete :-?
        break;
      case "gmail":
        domains = ["gmail-smtp-in.l.google.com"];
        break;
      case 'google':
        domains = ["aspmx.l.google.com", "googlemail.com", "gmail-smtp-in.l.google.com"];
        break;
      default:
        domains = ["aspmx.l.google.com", "googlemail.com", "gmail-smtp-in.l.google.com"];
    }

    const domainToCheck = email.includes("@") ? email.match(/.*@(.+)$/)[1] : email;

    return NSLookup('MX', domainToCheck, context).then(mxRecords => domains.some(domain => mxRecords.includes(domain)));

  }

});
