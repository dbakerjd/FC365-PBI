import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SeatsTableComponent } from './seats-table.component';

describe('SeatsTableComponent', () => {
  let component: SeatsTableComponent;
  let fixture: ComponentFixture<SeatsTableComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ SeatsTableComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(SeatsTableComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
