import { Pipe, PipeTransform } from '@angular/core';
@Pipe({ name: 'eventIDURL' })
export class EventIDURL implements PipeTransform {
    transform(msg: string) {
        if (msg)
            return msg.replace(/\s/gi, '+');
    }
}
