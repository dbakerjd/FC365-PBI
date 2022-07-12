import { ComponentFixture, TestBed } from '@angular/core/testing';

import { GeneralAreaComponent } from './general-area.component';

describe('GeneralAreaComponent', () => {
  let component: GeneralAreaComponent;
  let fixture: ComponentFixture<GeneralAreaComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ GeneralAreaComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(GeneralAreaComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
