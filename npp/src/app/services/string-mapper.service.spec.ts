import { TestBed } from '@angular/core/testing';

import { StringMapperService } from './string-mapper.service';

describe('StringMapperService', () => {
  let service: StringMapperService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(StringMapperService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
