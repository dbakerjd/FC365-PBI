import { TestBed } from '@angular/core/testing';

import { LicensingService } from './licensing.service';

describe('LicensingService', () => {
  let service: LicensingService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(LicensingService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
