import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/configMathTab1Tab/index.html")
@PreventIframe("/configMathTab1Tab/config.html")
@PreventIframe("/configMathTab1Tab/remove.html")
export class ConfigMathTab1Tab {
}
