import { Pipe, PipeTransform } from "@angular/core";
@Pipe({ name: "numformat" })
export class NumFormatPipe implements PipeTransform {
    transform(value: number, afterdot: number, alternate: string): string {
        if (value && afterdot) {
            if (afterdot === 0)
                return "0";
            return value.toFixed(afterdot);
        }
        else
            if (alternate)
                return alternate;
            else
                return "";
    }
}
