import { TestBed } from '@angular/core/testing';

import { SelectListsService } from './select-lists.service';

describe('SelectListsService', () => {
  let service: SelectListsService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(SelectListsService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
