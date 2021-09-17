import { ComponentFixture, TestBed } from '@angular/core/testing';

import { AuthEndComponent } from './auth-end.component';

describe('AuthEndComponent', () => {
  let component: AuthEndComponent;
  let fixture: ComponentFixture<AuthEndComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ AuthEndComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(AuthEndComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
