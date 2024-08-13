import { Component, OnInit } from "@angular/core";
import { DataService } from "../services/data.service";
import { IServer, DiskInfo, PingResult, RebootsInfo } from "../interfaces/interface";
import * as moment from "moment";

@Component({
    selector: "home",
    template: require("./home.component.html")
})

export class HomeComponent implements OnInit {
    servers: IServer[];
    ping: PingResult;
    noinfonum: number = 0;
    noinfo: string;
    srvnames: string[];

    constructor(private data: DataService) { }

    difference(a1: string[], a2: string[]) {
        let a2Set = new Set(a2);
        return a1.filter(function (x) { return !a2Set.has(x); });
    }

    noInfo(): string {
        if (this.servers && this.ping && this.srvnames) {
            let totalServers = this.ping.Dead.length + this.ping.Alive.length;
            let totalServersPingable = this.ping.Alive.length;
            let diff: string[] = (this.difference(this.ping.Alive, this.srvnames));
            this.noinfonum = diff.length;
            return diff.join(", ");
        }
    }

    ngOnInit() {

        this.noinfonum = 0;
        let names: string[] = [];

        this.data.getPingData().subscribe(
            data => this.ping = data,
            err => console.log(err),
            () => { }
        );

        this.data.getData().subscribe(
            data => this.servers = data,
            err => console.log(err),
            () => {
                this.servers = this.servers.filter(function (f) {
                    return f.OS != null;
                });
                // Completed

                let warnings24h: number = 0;
                let errors24h: number = 0;
                let reboots24h: number = 0;

                let now: moment.Moment = moment(); // NOW
                let time24Hago: moment.Moment = now.hours(now.hours() - 24);

                this.servers.forEach(function (srv) {
                    // Turn it to array
                    srv.Memory = [].concat(srv.Memory);
                    srv.CPU = [].concat(srv.CPU);

                    warnings24h = 0;
                    errors24h = 0;
                    reboots24h = 0;
                    srv.Warnings7d = 0;
                    srv.Errors7d = 0;
                    srv.Reboots7d = 0;

                    names.push(srv.SysName);

                    if (JSON.stringify(srv.Warnings) === "{}" || srv.Warnings === null) {
                        warnings24h = 0;
                        srv.Warnings7d = 0;
                    } else {
                        srv.Warnings = [].concat(srv.Warnings);
                        srv.Warnings7d = srv.Warnings.length;
                        srv.Warnings.forEach(function (w) {
                            let wtg: moment.Moment = moment(w.TimeGenerated, "DD-MMM-YYYY HH:mm:ss");
                            if (wtg > time24Hago)
                                warnings24h++;
                        });
                    }

                    if (JSON.stringify(srv.Errors) === "{}" || srv.Errors === null) {
                        errors24h = 0;
                        srv.Errors7d = 0;
                    } else {
                        srv.Errors = [].concat(srv.Errors);
                        srv.Errors7d = srv.Errors.length;
                        srv.Errors.forEach(function (e) {
                            let etg: moment.Moment = moment(e.TimeGenerated, "DD-MMM-YYYY HH:mm:ss");
                            if (etg > time24Hago)
                                errors24h++;
                        });
                    }

                    if (JSON.stringify(srv.Reboots) === "{}" || srv.Reboots === null || srv.Reboots === undefined) {
                        srv.Reboots7d = 0;
                        srv.Reboots24h = 0;
                    }
                    else {
                        srv.Reboots = [].concat(srv.Reboots);
                        srv.Reboots7d = srv.Reboots.length;
                        srv.Reboots.forEach(function (r) {
                            let rtg: moment.Moment = moment(r.TimeGenerated, "DD-MMM-YYYY HH:mm:ss");
                            if (rtg > time24Hago)
                                reboots24h++;
                        });
                    }

                    srv.Warnings24h = warnings24h;
                    srv.Errors24h = errors24h;
                    srv.Reboots24h = reboots24h;
                });
                this.srvnames = names;
            });
    }

    diskFreePercent(disk: DiskInfo) {
        return (disk.FreeSpace / disk.Size) * 100.0;
    }

    isArray(value: any) {
        return value instanceof Array;
    }

    networkIPFlatList(ips: any) {
        return ips.join(", ");
    }

    toArray(item: any) {
        return [].concat(item);
    }
}
