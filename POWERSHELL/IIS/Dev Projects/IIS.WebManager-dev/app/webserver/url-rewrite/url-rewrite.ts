﻿import { Metadata } from '../../common/metadata';

export class UrlRewrite {
    public id: string;
    public scope: string;
    public website: any;
    public _links: any;
}

export class AllowedServerVariablesSection {
    public id: string;
    public scope: string;
    public entries: Array<string>;
    public metadata: Metadata;
    public url_rewrite: UrlRewrite;
    public _links: any;
}

export class GlobalSection {
    public id: string;
    public scope: string;
    // Added in 2.1
    public use_original_url_encoding: boolean;
    public metadata: Metadata;
    public url_rewrite: UrlRewrite;
    public _links: any;
}

export class InboundSection {
    public id: string;
    public scope: string;
    // Added in 2.1
    public use_original_url_encoding: boolean;
    public metadata: Metadata;
    public url_rewrite: UrlRewrite;
    public _links: any;
}

export class OutboundSection {
    public id: string;
    public scope: string;
    public rewrite_before_cache: boolean;
    public metadata: Metadata;
    public url_rewrite: UrlRewrite;
    public _links: any;
}

export class ProvidersSection {
    public id: string;
    public scope: string;
    public metadata: Metadata;
    public url_rewrite: UrlRewrite;
    public _links: any;
}

export class RewriteMapsSection {
    public id: string;
    public scope: string;
    public metadata: Metadata;
    public url_rewrite: UrlRewrite;
    public _links: any;
}

export class InboundRule {
    constructor() {
        this.action = new Action();
    }

    public name: string;
    public id: string;
    public priority: number;
    public pattern: string;
    public pattern_syntax: PatternSyntax;
    public ignore_case: boolean;
    public negate: boolean;
    public stop_processing: boolean;
    public response_cache_directive: ResponseCacheDirective;
    public condition_match_constraints: ConditionMatchConstraints;
    public track_all_captures: boolean;
    public action: Action;
    public server_variables: Array<ServerVariableAssignment>;
    public conditions: Array<Condition>;
    public url_rewrite: UrlRewrite;
    public _links: any;

}

export class OutboundRule {
    public name: string;
    public id: string;
    public priority: number;
    public precondition: PreCondition;
    public match_type: OutboundMatchType;
    public server_variable: string;
    public tag_filters: OutboundTags;
    public pattern: string;
    public pattern_syntax: PatternSyntax;
    public ignore_case: boolean;
    public negate: boolean;
    public stop_processing: boolean;
    public enabled: boolean;
    public rewrite_value: string;
    public replace_server_variable: true;
    public condition_match_constraints: ConditionMatchConstraints;
    public track_all_captures: boolean;
    public conditions: Array<Condition>;
    public url_rewrite: UrlRewrite;
    public _links: any;
}

export class PreCondition {
    public name: string;
    public id: string;
    public match: MatchConstraint;
    public PatternSyntax: PatternSyntax;
    public requirements: Array<Condition>;
    public url_rewrite: UrlRewrite;
    public _links: any;
}

export class CustomTags {
    public name: string;
    public id: string;
    public tags: Array<CustomTag>;
    public url_rewrite: UrlRewrite;
    public _links: any;
}

export class Provider {
    public name: string;
    public id: string;
    public type: string;
    public settings: Array<ProviderSetting>;
    public url_rewrite: UrlRewrite;
    public _links: any;
}

export class RewriteMap {
    public name: string;
    public id: string;
    public default_value: string;
    public ignore_case: boolean;
    public mappings: Array<RewriteMapping>;
    public url_rewrite: UrlRewrite;
    public _links: any;
}

export class Action {
    public type: ActionType;
    public url: string;
    public append_query_string: boolean;
    public log_rewritten_url: boolean;
    public redirect_type: RedirectType;
    public status_code: number;
    public sub_status_code: number;
    public description: string;
    public reason: string;
}

export class ServerVariableAssignment {
    public name: string;
    public value: string;
    public replace: boolean;
}

export class Condition {
    public input: string;
    public pattern: string;
    public negate: boolean;
    public ignore_case: boolean;
    public match_type: MatchType;
}

export class OutboundTags {
    public a: boolean;
    public area: boolean;
    public base: boolean;
    public form: boolean;
    public frame: boolean;
    public head: boolean;
    public iframe: boolean;
    public img: boolean;
    public input: boolean;
    public link: boolean;
    public script: boolean;
    public custom: CustomTags;
}

export class CustomTag {
    public name: string;
    public attribute: string;
}

export class ProviderSetting {
    public name: string;
    public value: string;
}

export class RewriteMapping {
    public name: string;
    public value: string;
}

export type ActionType = "none" | "rewrite" | "redirect" | "custom_response" | "abort_request";
export const ActionType = {
    None: "none" as ActionType,
    Rewrite: "rewrite" as ActionType,
    Redirect: "redirect" as ActionType,
    CustomResponse: "custom_response" as ActionType,
    AbortRequest: "abort_request" as ActionType
}

export type RedirectType = "permanent" | "found" | "seeother" | "temporary";
export const RedirectType = {
    Permanent: "permanent" as RedirectType,
    Found: "found" as RedirectType,
    SeeOther: "seeother" as RedirectType,
    Temporary: "temporary" as RedirectType
}

export type ConditionMatchConstraints = "match_any" | "match_all";
export const ConditionMatchConstraints = {
    MatchAny: "match_any" as ConditionMatchConstraints,
    MatchAll: "match_all" as ConditionMatchConstraints
}

export type ResponseCacheDirective = "auto" | "always" | "never" | "not_if_rule_matched";
export const ResponseCacheDirective = {
    Auto: "auto" as ResponseCacheDirective,
    Always: "always" as ResponseCacheDirective,
    Never: "never" as ResponseCacheDirective,
    NotIfRuleMatched: "not_if_rule_matched" as ResponseCacheDirective
}

export type PatternSyntax = "regular_expression" | "wildcard" | "exact_match";
export const PatternSyntax = {
    RegularExpression: "regular_expression" as PatternSyntax,
    Wildcard: "wildcard" as PatternSyntax,
    ExactMatch: "exact_match" as PatternSyntax
}

export type MatchType = "pattern" | "is_file" | "is_directory";
export const MatchType = {
    Pattern: "pattern" as MatchType,
    IsFile: "is_file" as MatchType,
    IsDirectory: "is_directory" as MatchType
}

export type OutboundMatchType = "server_variable" | "response";
export const OutboundMatchType = {
    ServerVariable: "server_variable" as OutboundMatchType,
    Response: "response" as OutboundMatchType
}

export type MatchConstraint = "all" | "any";
export const MatchConstraint = {
    All: "all" as MatchConstraint,
    Any: "any" as MatchConstraint,
}

export class ActionTypeHelper {
    public static toFriendlyActionType(actionType: ActionType): string {
        switch (actionType) {
            case ActionType.AbortRequest:
                return "Abort";
            case ActionType.CustomResponse:
                return "Custom Response";
            case ActionType.Rewrite:
                return "Rewrite";
            case ActionType.Redirect:
                return "Redirect";
            case ActionType.None:
                return "None";
            default:
                throw new Error("Invalid action type");
        }
    }
}

export class OutboundMatchTypeHelper {
    public static display(matchType: OutboundMatchType): string {
        switch (matchType) {
            case OutboundMatchType.Response:
                return "Response";
            case OutboundMatchType.ServerVariable:
                return "Server Variable";
            default:
                throw new Error("Invalid match type");
        }
    }
}

export const IIS_SERVER_VARIABLES = [
    "ALL_HTTP",
    "ALL_RAW",
    "APP_POOL_ID",
    "APPL_MD_PATH",
    "APPL_PHYSICAL_PATH",
    "AUTH_PASSWORD",
    "AUTH_TYPE",
    "AUTH_USER",
    "CACHE_URL",
    "CERT_COOKIE",
    "CERT_FLAGS",
    "CERT_ISSUER",
    "CERT_KEYSIZE",
    "CERT_SECRETKEYSIZE",
    "CERT_SERIALNUMBER",
    "CERT_SERVER_ISSUER",
    "CERT_SERVER_SUBJECT",
    "CERT_SUBJECT",
    "CONTENT_LENGTH",
    "CONTENT_TYPE",
    "GATEWAY_INTERFACE",
    "HTTP_ACCEPT",
    "HTTP_ACCEPT_ENCODING",
    "HTTP_ACCEPT_LANGUAGE",
    "HTTP_CONNECTION",
    "HTTP_COOKIE",
    "HTTP_HOST",
    "HTTP_METHOD",
    "HTTP_REFERER",
    "HTTP_URL",
    "HTTP_USER_AGENT",
    "HTTP_VERSION",
    "HTTPS",
    "HTTPS_KEYSIZE",
    "HTTPS_SECRETKEYSIZE",
    "HTTP_SERVER_ISSUER",
    "HTTPS_SERVER_SUBJECT",
    "INSTANCE_ID",
    "INSTANCE_META_PATH",
    "LOCAL_ADDR",
    "LOGON_USER",
    "PATH_INFO",
    "PATH_TRANSLATE",
    "QUERY_STRING",
    "REMOTE_ADDR",
    "REMOTE_HOST",
    "REMOTE_PORT",
    "REMOTE_USER",
    "REQUEST_METHOD",
    "SCRIPT_NAME",
    "SCRIPT_TRANSLATED",
    "SERVER_NAME",
    "SERVER_PORT",
    "SERVER_PORT_SECURE",
    "SERVER_PROTOCOL",
    "SERVER_SOFTWARE",
    "SSI_EXEC_DISABLED",
    "UNENCODED_URL",
    "UNMAPPED_REMOTE_USER",
    "URL",
    "URL_PATH_INFO"
]