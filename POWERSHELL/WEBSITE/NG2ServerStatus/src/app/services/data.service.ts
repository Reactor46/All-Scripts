import { Injectable } from "@angular/core";
import { Http, Response, Headers, RequestOptions } from "@angular/http";
import { IServer, PingResult } from "../interfaces/interface";
import { Observable } from "rxjs/Rx";

@Injectable()
export class DataService {
    constructor(private http: Http) { }

    private getHeaders() {
        let headers = new Headers();
        headers.append("Accept", "application/json");
        return headers;
    }

    getData(): Observable<IServer[]> {
        return this.http.get("assets/servers.json").map(res => res.json());
    }

    getPingData(): Observable<PingResult> {
        return this.http.get("assets/ping.json").map(res => res.json());
    }
}
