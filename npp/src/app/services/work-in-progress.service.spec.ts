import { TestBed } from '@angular/core/testing';

import { WorkInProgressService } from './work-in-progress.service';

describe('WorkInProgressService', () => {
  let service: WorkInProgressService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(WorkInProgressService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
