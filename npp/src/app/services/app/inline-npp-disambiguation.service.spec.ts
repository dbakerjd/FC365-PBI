import { TestBed } from '@angular/core/testing';

import { InlineNppDisambiguationService } from './inline-npp-disambiguation.service';

describe('InlineNppDisambiguationService', () => {
  let service: InlineNppDisambiguationService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(InlineNppDisambiguationService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
